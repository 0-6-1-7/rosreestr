function LinksList() {
  try {
    const query = 'from:noreply-site@rosreestr.ru is:unread subject:Заявление с номером OfSite-* исполнено';
    const threads = GmailApp.search(query);

    if (threads.length === 0) {
      console.log('Нет подходящих сообщений');
    } else {
      let result = '';
      let myEmail = '';
      let linkCounter = 1;
      for (thread of threads) {
        const messages = thread.getMessages();
        for (message of messages) {
          const messageBody = message.getBody();
          const messageSubj = message.getSubject()
          const link = getLinkFromBody(messageBody);
          myEmail = message.getTo()
          result += String(`${linkCounter}.&nbsp;<a href="${link}">${messageSubj}</a><br>`);
          linkCounter ++;
          GmailApp.starMessage(message);
          GmailApp.markMessageRead(message);
        }
      }
      GmailApp.sendEmail(myEmail, 'Ссылки на загрузку результатов', '', {
        htmlBody: result
        })
    }
  } catch (err) {
    console.log(`При выполнении скрипта возникла ошибка: ${err.toString()}`);
  }
}

function getLinkFromBody(body) {
  const linkRegex = /<a\s+(?:[^>]*?\s+)?href="([^"]*)"/g;
  let links;
  let match;

  if ((match = linkRegex.exec(body)) !== null) {
    links = match[1];
  }

  return links;
}
