/**
 * BCN Comedy · website forms
 * Receives the forms from bcncomedy.com, logs them in a Google Sheet,
 * emails a branded notification to info@bcncomedy.com and a branded
 * confirmation to the person who filled in the form.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me (info@bcncomedy.com) · Who has access: Anyone
 */
const CFG = {
  INBOX: 'info@bcncomedy.com',
  FROM: 'info@bcncomedy.com',            // alias of the Workspace account running this script
  FROM_NAME: 'BCN Comedy',
  SITE: 'https://www.bcncomedy.com',
  INSTAGRAM: 'https://www.instagram.com/bcncomedy/',
  LOGO: 'https://www.bcncomedy.com/assets/email/bcn-logo.png',
  LASC: 'https://www.bcncomedy.com/assets/email/lasc-logo.png',
  SHEET_NAME: 'BCN Comedy · Website forms'
};

// Brand palette (same in light and dark mode on purpose)
const C = { bg:'#070909', card:'#131717', line:'#2a2f2f', ink:'#f4f1e6', muted:'#b9b6aa', sun:'#ECC122', sunInk:'#14120a', red:'#DF260E', lime:'#E6FB94' };

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });            // honeypot: bots fill it
    const type = ['alerts', 'book', 'contact'].indexOf(d.type) >= 0 ? d.type : 'contact';
    const email = String(d.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, error: 'email' });
    const lang = ['en', 'es', 'ca', 'pt'].indexOf(d.lang) >= 0 ? d.lang : 'en';
    const clean = {};
    Object.keys(d).forEach(k => { if (['type', 'website'].indexOf(k) < 0) clean[k] = String(d[k] == null ? '' : d[k]).slice(0, 4000); });

    logRow_(type, clean);
    notifyInbox_(type, clean);
    confirmVisitor_(type, lang, clean);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return json_({ ok: true, service: 'BCN Comedy forms' }); }

// Send from the info@ alias when Gmail has it under "Send mail as"; otherwise from the account itself.
function from_() { try { return GmailApp.getAliases().indexOf(CFG.FROM) >= 0 ? { from: CFG.FROM } : {}; } catch (e) { return {}; } }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Google Sheet log ---------- */
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(CFG.SHEET_NAME); props.setProperty('SHEET_ID', ss.getId()); }
  return ss;
}
const COLS = {
  alerts: ['Date', 'Name', 'Email', 'Language', 'Source'],
  book: ['Date', 'Name', 'Email', 'Company', 'Phone', 'Event date', 'Audience', 'Format', 'Show language', 'Message', 'Site language'],
  contact: ['Date', 'Name', 'Email', 'Message', 'Site language']
};
const KEYS = {
  alerts: ['name', 'email', 'lang', 'source'],
  book: ['name', 'email', 'company', 'phone', 'date', 'guests', 'format', 'language', 'message', 'lang'],
  contact: ['name', 'email', 'message', 'lang']
};
const TAB = { alerts: 'Show alerts', book: 'Bookings', contact: 'Messages' };
function logRow_(type, d) {
  const ss = sheet_();
  let sh = ss.getSheetByName(TAB[type]);
  if (!sh) {
    sh = ss.insertSheet(TAB[type]);
    sh.appendRow(COLS[type]); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS[type].length).setFontWeight('bold').setBackground(C.bg).setFontColor(C.sun);
    const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Hoja 1');
    if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  }
  sh.appendRow([new Date()].concat(KEYS[type].map(k => d[k] || '')));
}

/* ---------- emails ---------- */
const LBL = {
  name: 'Name', email: 'Email', company: 'Company', phone: 'Phone / WhatsApp', date: 'Event date', guests: 'Audience size',
  format: 'Format', language: 'Show language', message: 'Message', lang: 'Site language', source: 'Form'
};
const SUBJ = { alerts: 'New show-alerts signup', book: 'New booking request', contact: 'New message from the website' };

function notifyInbox_(type, d) {
  const rows = KEYS[type].filter(k => d[k]).map(k => [LBL[k] || k, d[k]]);
  const body =
    h1_(SUBJ[type]) +
    p_('Sent from the form on <a href="' + CFG.SITE + '" style="color:' + C.sun + ';text-decoration:none">bcncomedy.com</a>.') +
    table_(rows) +
    button_('mailto:' + encodeURIComponent(d.email) + '?subject=' + encodeURIComponent('Re: BCN Comedy'), 'Reply to ' + (d.name || d.email));
  GmailApp.sendEmail(CFG.INBOX, 'BCN Comedy · ' + SUBJ[type] + (d.name ? ' · ' + d.name : ''), plain_(rows), Object.assign({
    htmlBody: layout_(body, SUBJ[type]), name: 'BCN Comedy website', replyTo: d.email
  }, from_()));
}

const T = {
  en: {
    alerts: { s: "You're on the list! 🎤", h: "You're on the list", p: "Thanks{n}! We'll let you know about new dates, guest line-ups and special nights of the Laughin American Social Club." },
    book: { s: 'We got your request · BCN Comedy', h: 'Request received', p: "Thanks{n}! We've got your request and will get back to you within 48 hours with a proposal." },
    contact: { s: 'Thanks for your message · BCN Comedy', h: 'Message received', p: "Thanks{n}! We've got your message and will reply soon." },
    next: 'Next show', tickets: 'Get tickets', every: 'Live stand-up comedy in English · Thursdays 8:30PM · Bar AlgoRitmo, Gràcia',
    foot: 'You are receiving this email because you filled in a form on bcncomedy.com.', unsub: 'Unsubscribe: just reply “unsubscribe”.', follow: 'Follow us on Instagram'
  },
  es: {
    alerts: { s: '¡Ya estás en la lista! 🎤', h: 'Ya estás en la lista', p: '¡Gracias{n}! Te avisaremos de nuevas fechas, invitados y noches especiales del Laughin American Social Club.' },
    book: { s: 'Hemos recibido tu solicitud · BCN Comedy', h: 'Solicitud recibida', p: '¡Gracias{n}! Hemos recibido tu solicitud y te responderemos en 48 horas con una propuesta.' },
    contact: { s: 'Gracias por tu mensaje · BCN Comedy', h: 'Mensaje recibido', p: '¡Gracias{n}! Hemos recibido tu mensaje y te responderemos pronto.' },
    next: 'Próximo show', tickets: 'Comprar entradas', every: 'Stand-up en inglés · jueves 20:30 · Bar AlgoRitmo, Gràcia',
    foot: 'Recibes este email porque rellenaste un formulario en bcncomedy.com.', unsub: 'Para darte de baja, responde “baja”.', follow: 'Síguenos en Instagram'
  },
  ca: {
    alerts: { s: 'Ja ets a la llista! 🎤', h: 'Ja ets a la llista', p: 'Gràcies{n}! T’avisarem de noves dates, convidats i nits especials del Laughin American Social Club.' },
    book: { s: 'Hem rebut la teva sol·licitud · BCN Comedy', h: 'Sol·licitud rebuda', p: 'Gràcies{n}! Hem rebut la teva sol·licitud i et respondrem en 48 hores amb una proposta.' },
    contact: { s: 'Gràcies pel teu missatge · BCN Comedy', h: 'Missatge rebut', p: 'Gràcies{n}! Hem rebut el teu missatge i et respondrem aviat.' },
    next: 'Proper show', tickets: 'Compra entrades', every: 'Stand-up en anglès · dijous 20:30 · Bar AlgoRitmo, Gràcia',
    foot: 'Reps aquest email perquè vas omplir un formulari a bcncomedy.com.', unsub: 'Per donar-te de baixa, respon “baixa”.', follow: 'Segueix-nos a Instagram'
  },
  pt: {
    alerts: { s: 'Você está na lista! 🎤', h: 'Você está na lista', p: 'Valeu{n}! Vamos te avisar das novas datas, convidados e noites especiais do Laughin American Social Club.' },
    book: { s: 'Recebemos seu pedido · BCN Comedy', h: 'Pedido recebido', p: 'Valeu{n}! Recebemos seu pedido e respondemos em até 48 horas com uma proposta.' },
    contact: { s: 'Obrigado pela mensagem · BCN Comedy', h: 'Mensagem recebida', p: 'Valeu{n}! Recebemos sua mensagem e respondemos em breve.' },
    next: 'Próximo show', tickets: 'Garantir ingressos', every: 'Stand-up em inglês · quintas 20:30 · Bar AlgoRitmo, Gràcia',
    foot: 'Você recebeu este e-mail porque preencheu um formulário em bcncomedy.com.', unsub: 'Para sair da lista, responda “sair”.', follow: 'Siga no Instagram'
  }
};

function confirmVisitor_(type, lang, d) {
  const t = T[lang], m = t[type];
  const first = String(d.name || '').trim().split(/\s+/)[0];
  const n = first ? ', ' + esc_(first) : '';
  let body =
    '<tr><td align="center" style="padding:6px 28px 4px"><img src="' + CFG.LASC + '" width="260" alt="Laughin American Social Club" style="display:block;width:260px;max-width:80%;height:auto;border:0"></td></tr>' +
    h1_(m.h) + p_(m.p.replace('{n}', n));
  if (d.nextShow) {
    body +=
      '<tr><td style="padding:10px 28px 6px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:14px">' +
      '<tr><td style="padding:18px 20px;font-family:Arial,Helvetica,sans-serif">' +
      '<div style="font-size:12px;letter-spacing:2px;text-transform:uppercase;color:' + C.red + ';font-weight:bold">' + t.next + '</div>' +
      '<div style="font-size:22px;line-height:1.2;color:' + C.sun + ';font-weight:900;text-transform:uppercase;margin:6px 0 4px">' + esc_(d.nextShow) + '</div>' +
      '<div style="font-size:14px;color:' + C.muted + '">' + t.every + '</div></td></tr></table></td></tr>';
    if (d.nextUrl) body += button_(d.nextUrl, t.tickets);
  }
  body += '<tr><td align="center" style="padding:14px 28px 4px;font-family:Arial,Helvetica,sans-serif;font-size:14px"><a href="' + CFG.INSTAGRAM + '" style="color:' + C.sun + ';text-decoration:none;font-weight:bold">' + t.follow + ' → @bcncomedy</a></td></tr>';
  const foot = t.foot + (type === 'alerts' ? ' ' + t.unsub : '');
  GmailApp.sendEmail(d.email, m.s, m.h + '\n\n' + m.p.replace('{n}', first ? ', ' + first : '') + '\n\n' + CFG.SITE, Object.assign({
    htmlBody: layout_(body, m.s, foot), name: CFG.FROM_NAME, replyTo: CFG.INBOX
  }, from_()));
}

/* ---------- HTML building blocks ---------- */
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function h1_(s) { return '<tr><td style="padding:18px 28px 6px;font-family:Arial Black,Arial,Helvetica,sans-serif;font-size:26px;line-height:1.15;font-weight:900;text-transform:uppercase;color:' + C.ink + '">' + esc_(s) + '</td></tr>'; }
function p_(html) { return '<tr><td style="padding:6px 28px 10px;font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.55;color:' + C.muted + '">' + html + '</td></tr>'; }
function button_(href, label) {
  return '<tr><td align="center" style="padding:14px 28px 8px"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td style="border-radius:999px;background:' + C.sun + ';box-shadow:0 4px 0 #9c7e10">' +
    '<a href="' + href + '" style="display:inline-block;padding:14px 28px;font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:bold;color:' + C.sunInk + ';text-decoration:none;border-radius:999px">' + esc_(label) + '</a></td></tr></table></td></tr>';
}
function table_(rows) {
  return '<tr><td style="padding:8px 28px 6px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:14px">' +
    rows.map((r, i) => '<tr><td style="padding:12px 16px;' + (i ? 'border-top:1px solid ' + C.line + ';' : '') + 'font-family:Arial,Helvetica,sans-serif;font-size:12px;letter-spacing:1px;text-transform:uppercase;color:' + C.sun + ';font-weight:bold;width:34%;vertical-align:top">' + esc_(r[0]) +
      '</td><td style="padding:12px 16px;' + (i ? 'border-top:1px solid ' + C.line + ';' : '') + 'font-family:Arial,Helvetica,sans-serif;font-size:15px;color:' + C.ink + ';white-space:pre-wrap">' + esc_(r[1]) + '</td></tr>').join('') +
    '</table></td></tr>';
}
function plain_(rows) { return rows.map(r => r[0] + ': ' + r[1]).join('\n'); }
function layout_(inner, preheader, foot) {
  // Dark brand design, declared for both colour schemes so mail apps keep the same colours.
  return '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="color-scheme" content="light dark"><meta name="supported-color-schemes" content="light dark">' +
    '<style>:root{color-scheme:light dark;supported-color-schemes:light dark}' +
    'body,.bg{background:' + C.bg + '!important}' +
    '@media (prefers-color-scheme:dark){body,.bg{background:' + C.bg + '!important}.card{background:' + C.bg + '!important}}' +
    '[data-ogsc] .bg,[data-ogsb] .bg{background:' + C.bg + '!important}' +
    'a{color:' + C.sun + '}</style></head>' +
    '<body class="bg" style="margin:0;padding:0;background:' + C.bg + '">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0">' + esc_(preheader || '') + '</div>' +
    '<table role="presentation" class="bg" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.bg + '" bgcolor="' + C.bg + '"><tr><td align="center" style="padding:24px 12px">' +
    '<table role="presentation" class="card" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:' + C.bg + ';border:1px solid ' + C.line + ';border-radius:18px;overflow:hidden" bgcolor="' + C.bg + '">' +
    '<tr><td style="height:8px;background:' + C.red + ';background-image:radial-gradient(circle,' + C.sun + ' 2px,transparent 2.5px);background-size:12px 8px" bgcolor="' + C.red + '"></td></tr>' +
    '<tr><td align="center" style="padding:26px 28px 6px"><a href="' + CFG.SITE + '"><img src="' + CFG.LOGO + '" width="150" alt="BCN Comedy" style="display:block;width:150px;height:auto;border:0"></a></td></tr>' +
    inner +
    '<tr><td style="padding:22px 28px 26px;font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:1.5;color:#8a877d;border-top:1px solid ' + C.line + '">' +
    (foot ? esc_(foot) + '<br>' : '') + 'BCN Comedy · Barcelona · <a href="' + CFG.SITE + '" style="color:' + C.sun + ';text-decoration:none">bcncomedy.com</a> · <a href="mailto:' + CFG.INBOX + '" style="color:' + C.sun + ';text-decoration:none">' + CFG.INBOX + '</a></td></tr>' +
    '</table></td></tr></table></body></html>';
}

/* Run once from the editor to authorize and send yourself a preview of every email. */
function checkAlias() { console.log('Aliases:', GmailApp.getAliases()); console.log('Sending from:', from_().from || Session.getActiveUser().getEmail()); }

function testAll() {
  const base = { name: 'Test Person', email: CFG.INBOX, lang: 'en', nextShow: 'Thursday 3 December · 8:30PM', nextUrl: CFG.SITE + '/agenda' };
  ['alerts', 'book', 'contact'].forEach(type => {
    const d = Object.assign({}, base, type === 'book' ? { company: 'ACME', phone: '+34 600 000 000', date: '2026-12-20', guests: '60', format: 'Private comedy night', language: 'English', message: 'Test booking request' } :
      type === 'contact' ? { message: 'Test message' } : { source: 'test' });
    logRow_(type, d); notifyInbox_(type, d); confirmVisitor_(type, 'en', d);
  });
}
