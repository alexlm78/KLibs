package com.alexlm78.klibs.mail;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Iterator;
import java.util.List;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

public class KMail {
    private final ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

    public KMail( String username, String password )throws URISyntaxException {
        this.service.setCredentials(new WebCredentials(username, password));
        this.service.setUrl(new URI("https://outlook.office365.com/owa/claro.com.gt/ews/exchange.asmx"));
    }

    public KMail() throws URISyntaxException {
        this("controltareas.pisa", "claro+01");
    }

    public Boolean sengMail(String subject, String message, List<String> recipients, List<String> filesNames){
        try {
            EmailMessage email = new EmailMessage(this.service);
            email.setSubject(subject);
            email.setBody(new MessageBody(message));

            Iterator<String> localIterator;
            String fileName;
            for ( localIterator = filesNames.iterator(); localIterator.hasNext(); email.getAttachments()
                    .addFileAttachment(fileName)) {
                fileName = (String) localIterator.next();
            }
            String recipient;
            for (localIterator = recipients.iterator(); localIterator.hasNext(); email.getToRecipients()
                    .add(recipient)) {
                recipient = (String) localIterator.next();
            }
            email.sendAndSaveCopy();
            return true;
        } catch (Exception ex) {
            ex.getMessage();
            ex.printStackTrace();
            return false;
        }
    }
}
