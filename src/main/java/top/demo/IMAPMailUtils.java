package top.demo;

import javax.mail.*;
import java.io.IOException;
import java.util.Properties;

public class IMAPMailUtils {
    //配置类
    private static final Properties properties;

    static {
        properties = new Properties();
        try {
            properties.load(IMAPMailUtils.class.getClassLoader().getResourceAsStream("config.properties"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void delMail() {
        Session session = Session.getInstance(properties);
        Store store = null;
        Folder folder = null;
        String mailUsername = properties.getProperty("mail_username");
        String mailPwd = properties.getProperty("mail_password");
        String protocol = properties.getProperty("mail.store.protocol");
        boolean flag = false;
        try {
            store = session.getStore(protocol);
            store.connect(mailUsername, mailPwd);
            folder = store.getFolder("草稿夹");
            folder.open(Folder.READ_WRITE);
            Message[] messages = folder.getMessages();
            for (Message message : messages) {
                if (message.getFlags().contains(Flags.Flag.SEEN)) {
                    message.setFlag(Flags.Flag.DELETED, true);
                    flag = true;
                }
            }
        } catch (MessagingException e) {
            e.printStackTrace();
        } finally {
            try {
                if (folder != null) {
                    if (flag) {
                        folder.close(true);
                    } else {
                        folder.close();
                    }
                }
                if (store != null) {
                    store.close();
                }
            } catch (MessagingException e) {
                e.printStackTrace();
            }
        }
    }
}
