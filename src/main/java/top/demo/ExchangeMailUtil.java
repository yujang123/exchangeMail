package top.demo;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Properties;

public class ExchangeMailUtil {

    private static final Properties PROPERTIES;

    static {
        PROPERTIES = new Properties();
        InputStreamReader isr = null;
        try {
            // 设置编码,防止中文乱码
            isr = new InputStreamReader(ExchangeMailUtil.class.getClassLoader().getResourceAsStream("config.properties"), "UTF-8");
            PROPERTIES.load(isr);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (isr != null) {
                try {
                    isr.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void main(String[] args) {
        // 生成可执行jar后参数绑定
        String userName;
        String userPwd;
        String folderName;
        String isRead = "";
        if (args.length >= 4) {
            userName = args[0];
            userPwd = args[1];
            folderName = args[2];
            isRead = args[3];
        } else {
            userName = args[0];
            userPwd = args[1];
            folderName = args[2];
        }
        delEmail(userName, userPwd, folderName, isRead);
    }

    /**
     * 删除指定文件夹下的邮件
     *
     * @param userName   登录用户名
     * @param userPwd    登陆密码
     * @param folderName 指定文件夹
     * @param isRead     邮件是否已读 false:未读邮件 true:已读邮件 不填:所有邮件
     * @return 是否删除成功
     */
    private static boolean delEmail(String userName, String userPwd, String folderName, String isRead) {
        boolean flag = false;
        // 从配置文件中获取参数
        // String exchangeUser = properties.getProperty("exchange.user");
        // String exchangePwd = properties.getProperty("exchange.pwd");
        String exchangeUser = userName;
        String exchangePwd = userPwd;
        String exchangeDomain = PROPERTIES.getProperty("exchange.domain");
        String exchangeUrl = PROPERTIES.getProperty("exchange.url");
        // String exchangeFolderName = properties.getProperty("exchange.folder.name");
        // String exchangeIsRead = properties.getProperty("exchange.isRead");
        String exchangeFolderName = folderName;
        String exchangeIsRead = isRead;

        // 实例化ExchangeService,并设置ExchangeService版本,用户名\密码\域认证信息
        ExchangeService exchangeService = new ExchangeService(ExchangeVersion.Exchange2010);
        ExchangeCredentials exchangeCredentials = new WebCredentials(exchangeUser, exchangePwd, exchangeDomain);
        exchangeService.setCredentials(exchangeCredentials);

        try {
            // 自动发现请求URL,请求速度极慢,所以断点抓出请求地址就可以注释掉了
            // exchangeService.autodiscoverUrl(properties.getProperty("exchange.address"));
            // 手动设置请求URL
            exchangeService.setUrl(new URI(exchangeUrl));
            // 搜索条件,邮件是否已读
            SearchFilter searchFilter = null;
            if (!exchangeIsRead.isEmpty()) {
                System.out.println("-------进来-------");
                searchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, exchangeIsRead);
            }
            // 控制获取数
            ItemView itemView = new ItemView(Integer.MAX_VALUE);
            FolderView folderView = new FolderView(Integer.MAX_VALUE);

            // 获取所有消息文件夹,遍历,存在指定文件夹则获取文件夹中的邮件并删除
            FindFoldersResults folders = exchangeService.findFolders(WellKnownFolderName.MsgFolderRoot, folderView);
            FindItemsResults<Item> items = null;

            for (Folder folder : folders.getFolders()) {
                if (folder.getDisplayName().equals(exchangeFolderName)) {
                    if (searchFilter != null) { // 根据条件获取Item
                        System.out.println("删除一部分");
                        items = folder.findItems(searchFilter, itemView);
                        break;
                    } else {    //获取所有Item
                        System.out.println("删除所有");
                        items = folder.findItems(itemView);
                        break;
                    }
                }
            }

            if (items != null && items.getTotalCount() > 0) {
                for (Item item : items) {
                    item.delete(DeleteMode.MoveToDeletedItems);
                }
                flag = true;
            }
        } catch (ServiceLocalException | URISyntaxException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return flag;
    }
}
