package com.garwen.delmail.task;

import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.net.URI;

/**
 *Email Task for Exchange mails
 *@author Garwen
 *@date 2020-2-17 0:19
 */

@Component
@EnableScheduling
@Slf4j
public class ScheduledEmailTask {
//    @Scheduled(cron = "0/5 * * * * ?")
//    private void testTask(){
//        log.info("Print this Message " + new Date().toString());
//        log.info("========================");
//    }

    /**
     *delete All the mails in the DeletedItems
     * the frequency is Every month 1st 23:30
     *@author Garwen
     *@date 2020-02-17 0:20
     *@param
     *@return void
     *@throws
     */
    @Scheduled(cron = "0 30 23 1 * ?")
    private void deleteEmail(){
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials("username", "password","domain");
        service.setCredentials(credentials);

        int size = 10;
        int offset = size;
        ItemView view = new ItemView(size);
        FindItemsResults<Item> findResults;

        try {
            service.setUrl(new URI("https://mail.crc.com.hk/ews/exchange.asmx"));
            do{
                findResults = service.findItems(WellKnownFolderName.DeletedItems, view);
                for( Item item: findResults.getItems()){
                    item.delete(DeleteMode.HardDelete);
//                    log.info("id: {}; sub: {}", item.getId(), item.getSubject());
                }
                offset += size;
                view.setOffset(offset);
            }while(findResults.isMoreAvailable());

        } catch (Exception e) {
            e.printStackTrace();
        }


    }
}
