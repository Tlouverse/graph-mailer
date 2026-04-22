const { GraphMailer } = require("@tlouverse/graph-mailer");

const mailer = new GraphMailer({
  "clientId": "",
  "tenantId": "",
  "clientSecret": ""
})

mailer.send({
  "to": ["thomas.louvet@akkodis.com"],
  "from": "noreply.portail.avv@akkodis.com",
  "subject": "Test du package graphmailer",
  "text": "Un simple texte devrait suffire pour faire ce premier test"
})