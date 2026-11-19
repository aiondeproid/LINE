/**
 * LINE公式アカウント 新着メッセージ通知 (Google Apps Script)
 *
 * 公式アカウントにメッセージが届いたら、次の2つを行う:
 *   1. 指定したアドレスにメールを送る
 *   2. 管理者のLINEにプッシュ通知を送る
 *
 * スクリプトプロパティ（プロジェクトの設定 > スクリプト プロパティ）:
 *   CHANNEL_ACCESS_TOKEN : チャネルアクセストークン（長期）            [必須]
 *   MAIL_TO              : 通知先メールアドレス（カンマ区切りで複数指定可）   [任意]
 *   ADMIN_USER_IDS       : 通知を受け取る管理者のuserId（カンマ区切り）     [任意]
 *   REPORT_KEYWORDS      : 体調不良報告を始めるキーワード（カンマ区切り）    [任意]
 *                          未設定なら「体調不良報告,体調不良,報告」
 *
 * 体調不良報告: ユーザーがキーワードを送ると、質問を1つずつ返信して回答を集める。
 * ボタン（クイックリプライ）で答えられる質問はボタンを出す。最後に確認して「送信」で通知する。
 * 回答の途中経過はスクリプトプロパティに「SESSION_<userId>」として一時保存する（24時間で無効）。
 *
 * 管理者のuserIdを調べる方法: 管理者が公式アカウントに「ID確認」と送る
 * → 本人のuserIdが返信される（返信メッセージは月の通数にカウントされない）。
 */

const PROPS = PropertiesService.getScriptProperties();
const TOKEN = PROPS.getProperty('CHANNEL_ACCESS_TOKEN');
const MAIL_TO = (PROPS.getProperty('MAIL_TO') || '').trim();
const ADMIN_IDS = (PROPS.getProperty('ADMIN_USER_IDS') || '')
  .split(',').map(s => s.trim()).filter(Boolean);

const ID_COMMAND = 'ID確認';
const REPORT_KEYWORDS = (PROPS.getProperty('REPORT_KEYWORDS') || '体調不良報告,体調不良,報告')
  .split(',').map(s => s.trim()).filter(Boolean);

// ===== 体調不良報告の質問 =====
// choices: ボタンで出す選択肢 / free: 選択肢以外の入力も受け付ける / picker: 日付・日時の選択ボタンを出す
const QUESTIONS = [
  { key: 'name', label: 'お名前', text: 'お名前を教えてください。' },
  { key: 'who', label: '本人／家族', text: 'ご本人ですか、ご家族ですか？', choices: ['本人', '家族'] },
  { key: 'when', label: '発生日時', text: '事象の発生日時はいつですか？\nボタンで選ぶか、「11/19 14:30」のように入力してください。', picker: 'datetime' },
  { key: 'where', label: '発生場所', text: '発生場所はどこですか？\nラインの場合はライン名を入力してください。', choices: ['トイレ', '更衣室', '休憩室', '家'], free: true },
  { key: 'type', label: '嘔吐・下痢', text: '嘔吐・下痢のどちらですか？', choices: ['嘔吐', '下痢', '両方'] },
  { key: 'food', label: '食品に触れる作業中', text: '直接食品に触れる作業中でしたか？', choices: ['はい', 'いいえ'] },
  { key: 'since', label: '自覚症状はいつごろから', text: '自覚症状はいつごろからありましたか？\nボタンで選ぶか、「昨日の夜から」のように入力してください。', picker: 'datetime' },
  { key: 'stoolCo', label: '発症当日の排便（会社トイレ）', text: '発症当日、会社のトイレで排便しましたか？', choices: ['した', 'しない'] },
  { key: 'stoolHome', label: '発症当日の排便（自宅トイレ）', text: '発症当日、自宅のトイレで排便しましたか？', choices: ['した', 'しない'] },
  { key: 'vomitCo', label: '発症当日の嘔吐（会社トイレ）', text: '発症当日、会社のトイレで嘔吐しましたか？', choices: ['した', 'しない'] },
  { key: 'vomitHome', label: '発症当日の嘔吐（自宅トイレ）', text: '発症当日、自宅のトイレで嘔吐しましたか？', choices: ['した', 'しない'] },
  { key: 'onset', label: '発症日', text: '発症日はいつですか？', picker: 'date' },
  { key: 'workPrev', label: '発症前日の作業内容', text: '自覚症状が出る前日の作業内容を教えてください。' },
  { key: 'workOnset', label: '発症日の作業内容', text: '発症日の作業内容を教えてください。' },
  { key: 'other', label: 'その他', text: 'その他、伝えたいことがあれば入力してください。', choices: ['なし'], free: true },
];

const CMD_BACK = '戻る';
const CMD_CANCEL = '中止';
const CMD_SUBMIT = '送信';
const CMD_RESTART = '最初からやり直す';
const SESSION_PREFIX = 'SESSION_';
const SESSION_TTL_MS = 24 * 60 * 60 * 1000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 連続送信で回答の保存がぶつからないように1件ずつ処理する
    const events = JSON.parse(e.postData.contents).events || [];
    events.forEach(handleEvent_);
  } catch (err) {
    console.error(err.stack || err);
  } finally {
    lock.releaseLock();
  }
  // 何があっても200を返す（LINE側の再送や無効化を防ぐ）
  return ContentService.createTextOutput('OK');
}

function handleEvent_(ev) {
  const userId = ev.source && ev.source.userId;

  // 日付・日時の選択ボタンの結果
  if (ev.type === 'postback') {
    const session = userId && getSession_(userId);
    const p = ev.postback.params || {};
    if (session) handleAnswer_(ev, userId, session, { picked: p.datetime || p.date || p.time });
    return;
  }
  if (ev.type !== 'message') return;

  const text = ev.message.type === 'text' ? ev.message.text.trim() : null;

  // 管理者のuserId確認用コマンド
  if (text === ID_COMMAND) {
    reply_(ev.replyToken, 'あなたのuserId:\n' + userId);
    return;
  }

  // 体調不良報告を始める（途中でもう一度送ったら最初から）
  if (text && REPORT_KEYWORDS.includes(text) && userId) {
    const session = { step: 0, answers: {}, startedAt: Date.now() };
    saveSession_(userId, session);
    replyMessages_(ev.replyToken, [
      textMsg_('体調不良報告を始めます。質問に1つずつ答えてください。\n（「戻る」で1つ前へ、「中止」でやめられます）'),
      questionMsg_(session),
    ]);
    return;
  }

  // 報告の途中なら回答として扱う（個別の通知はしない）
  const session = userId && getSession_(userId);
  if (session) {
    handleAnswer_(ev, userId, session, { text: text });
    return;
  }

  // 管理者自身のメッセージは通知しない（ループ防止）
  if (ADMIN_IDS.includes(userId)) return;

  const name = userId ? getDisplayName_(userId) : '(不明)';
  const time = Utilities.formatDate(new Date(ev.timestamp), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const body = [
    '【LINE公式アカウント 新着メッセージ】',
    '日時: ' + time,
    '送信者: ' + name,
    '内容: ' + describeMessage_(ev.message),
  ].join('\n');
  notify_('[LINE] ' + name + 'さんからメッセージ', body);
}

// ===== 体調不良報告: 回答の処理 =====

function handleAnswer_(ev, userId, s, input) {
  const text = input.text;

  if (text === CMD_CANCEL) {
    clearSession_(userId);
    reply_(ev.replyToken, '体調不良報告を中止しました。もう一度始めるときは「' + REPORT_KEYWORDS[0] + '」と送ってください。');
    return;
  }
  if (text === CMD_BACK) {
    s.step = Math.max(0, s.step - 1);
    saveSession_(userId, s);
    replyMessages_(ev.replyToken, [questionMsg_(s)]);
    return;
  }

  // 最後の確認画面
  if (s.step >= QUESTIONS.length) {
    if (text === CMD_SUBMIT) {
      submitReport_(userId, s);
      clearSession_(userId);
      reply_(ev.replyToken, '体調不良報告を受け付けました。担当者から連絡があるまでお待ちください。');
    } else if (text === CMD_RESTART) {
      s.step = 0;
      s.answers = {};
      saveSession_(userId, s);
      replyMessages_(ev.replyToken, [questionMsg_(s)]);
    } else {
      replyMessages_(ev.replyToken, [textMsg_('下のボタンから選んでください。'), confirmMsg_(s)]);
    }
    return;
  }

  const q = QUESTIONS[s.step];
  let value = null;
  if (input.picked && q.picker) {
    value = formatPicked_(input.picked);
  } else if (text) {
    const strict = q.choices && !q.free && !q.picker;
    if (!strict || q.choices.includes(text)) value = text;
  }

  if (value === null) {
    const hint = q.choices && !q.free ? '下のボタンから選んでください。' : '文字で入力してください。';
    replyMessages_(ev.replyToken, [textMsg_(hint), questionMsg_(s)]);
    return;
  }

  s.answers[q.key] = value;
  s.step++;
  saveSession_(userId, s);
  replyMessages_(ev.replyToken, [s.step < QUESTIONS.length ? questionMsg_(s) : confirmMsg_(s)]);
}

function submitReport_(userId, s) {
  const lineName = getDisplayName_(userId);
  const name = s.answers.name || lineName;
  const time = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const body = [
    '【LINE公式アカウント 体調不良報告】',
    '受付日時: ' + time,
    'LINE表示名: ' + lineName,
    '',
    summary_(s),
  ].join('\n');
  notify_('[LINE][体調不良報告] ' + name + 'さんから報告', body);
}

function summary_(s) {
  return QUESTIONS.map(q => '・' + q.label + '：' + (s.answers[q.key] || '')).join('\n');
}

function formatPicked_(v) {
  // "2026-11-19T14:30" / "2026-11-19" / "14:30"
  const m = v.match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}:\d{2}))?$/);
  if (!m) return v;
  return Number(m[2]) + '月' + Number(m[3]) + '日' + (m[4] ? ' ' + m[4] : '');
}

// ===== 体調不良報告: 送るメッセージ =====

function questionMsg_(s) {
  const q = QUESTIONS[s.step];
  const items = [];
  (q.choices || []).forEach(c => items.push(messageItem_(c)));
  if (q.picker) {
    items.push({
      type: 'action',
      action: {
        type: 'datetimepicker',
        label: q.picker === 'date' ? '日付を選ぶ' : '日時を選ぶ',
        data: 'picker=' + q.key,
        mode: q.picker,
      },
    });
  }
  if (s.step > 0) items.push(messageItem_(CMD_BACK));
  items.push(messageItem_(CMD_CANCEL));
  return withQuickReply_('(' + (s.step + 1) + '/' + QUESTIONS.length + ') ' + q.text, items);
}

function confirmMsg_(s) {
  return withQuickReply_(
    '以下の内容で送信します。よろしければ「送信」を押してください。\n\n' + summary_(s),
    [CMD_SUBMIT, CMD_BACK, CMD_RESTART, CMD_CANCEL].map(messageItem_)
  );
}

function messageItem_(label) {
  return { type: 'action', action: { type: 'message', label: label, text: label } };
}

function withQuickReply_(text, items) {
  return { type: 'text', text: text, quickReply: { items: items.slice(0, 13) } };
}

function textMsg_(text) {
  return { type: 'text', text: text };
}

// ===== 体調不良報告: 途中経過の保存 =====

function getSession_(userId) {
  const raw = PROPS.getProperty(SESSION_PREFIX + userId);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (Date.now() - s.startedAt > SESSION_TTL_MS) {
    clearSession_(userId);
    return null;
  }
  return s;
}

function saveSession_(userId, s) {
  PROPS.setProperty(SESSION_PREFIX + userId, JSON.stringify(s));
}

function clearSession_(userId) {
  PROPS.deleteProperty(SESSION_PREFIX + userId);
}

// ===== 共通 =====

function notify_(subject, body) {
  if (MAIL_TO) {
    try {
      MailApp.sendEmail(MAIL_TO, subject, body);
    } catch (err) {
      console.error('メール送信失敗: ' + err);
    }
  }
  if (ADMIN_IDS.length) {
    try {
      // 返信はOfficial Account Managerのチャット画面から行う
      multicast_(ADMIN_IDS, body + '\n\n※返信はLINE公式アカウントアプリのチャットから');
    } catch (err) {
      console.error('LINE通知失敗: ' + err);
    }
  }
}

function describeMessage_(m) {
  switch (m.type) {
    case 'text': return m.text;
    case 'image': return '[画像]';
    case 'video': return '[動画]';
    case 'audio': return '[音声]';
    case 'file': return '[ファイル] ' + (m.fileName || '');
    case 'location': return '[位置情報] ' + (m.address || m.title || '');
    case 'sticker': return '[スタンプ]';
    default: return '[' + m.type + ']';
  }
}

function getDisplayName_(userId) {
  try {
    const res = api_('get', '/v2/bot/profile/' + userId);
    return JSON.parse(res.getContentText()).displayName;
  } catch (err) {
    return userId; // ブロック済み・テスト用ID等で取得できない場合
  }
}

function reply_(replyToken, text) {
  replyMessages_(replyToken, [textMsg_(text)]);
}

function replyMessages_(replyToken, messages) {
  if (!replyToken) {
    // エディタからのテスト実行時は返信先がないのでログに出すだけ
    messages.forEach(m => console.log('[返信] ' + m.text));
    return;
  }
  try {
    api_('post', '/v2/bot/message/reply', { replyToken: replyToken, messages: messages });
  } catch (err) {
    console.error('返信失敗: ' + err);
  }
}

function multicast_(to, text) {
  api_('post', '/v2/bot/message/multicast', {
    to: to,
    messages: [{ type: 'text', text: text.slice(0, 5000) }],
  });
}

function api_(method, path, payload) {
  const options = {
    method: method,
    headers: { Authorization: 'Bearer ' + TOKEN },
    muteHttpExceptions: true,
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  const res = UrlFetchApp.fetch('https://api.line.me' + path, options);
  if (res.getResponseCode() >= 300) {
    throw new Error(path + ' ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  return res;
}

// ===== テスト（エディタから手動実行） =====

/** メールとLINE通知が届くか確認する */
function testNotify() {
  handleEvent_({
    type: 'message',
    timestamp: Date.now(),
    source: { type: 'user', userId: 'Utest' },
    message: { type: 'text', text: 'テスト通知です' },
  });
}

/** 体調不良報告を最初から最後まで答えた想定で流す。返信内容は実行ログに出て、最後に通知が届く */
function testReportFlow() {
  const send = (text) => handleEvent_({
    type: 'message',
    timestamp: Date.now(),
    source: { type: 'user', userId: 'Utest' },
    message: { type: 'text', text: text },
  });
  const pick = (value) => handleEvent_({
    type: 'postback',
    timestamp: Date.now(),
    source: { type: 'user', userId: 'Utest' },
    postback: { data: 'picker', params: value.includes('T') ? { datetime: value } : { date: value } },
  });

  send(REPORT_KEYWORDS[0]);
  send('テスト太郎');
  send('本人');
  pick('2026-11-19T14:30');
  send('第2ライン');
  send('両方');
  send('はい');
  send('昨日の夜から');
  send('した');
  send('しない');
  send('しない');
  send('した');
  pick('2026-11-19');
  send('仕込み');
  send('盛り付け');
  send('なし');
  send(CMD_SUBMIT);
}
