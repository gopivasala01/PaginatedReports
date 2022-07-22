package Package1;
import java.io.File;
import java.security.KeyStore.PasswordProtection;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.FluentWait;

public class SendAttachmentInMail 
{
   public void sendFileToMail(String agentName, String agentEmail,String managerEmail) 
   {
      // Recipient's email ID needs to be mentioned.
      String to = "bireports@beetlerim.com";

      // Sender's email ID needs to be mentioned
      String from = "gopi.v@beetlerim.com";

      final String username = "gopi.v@beetlerim.com";//change accordingly
      final String password = "Welcome@123";//change accordingly

      // Assuming you are sending email through relay.jangosmtp.net
      String host = "smtpout.asia.secureserver.net";

      Properties props = new Properties();
      props.put("mail.smtp.auth", "true");
      //props.put("mail.smtp.starttls.enable", "true");
     props.put("mail.smtp.host", host);
      props.put("mail.smtp.port", "80");

      // Get the Session object.
      Session session = Session.getInstance(props,
         new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
               return new PasswordAuthentication(AppConfig.fromEmail, AppConfig.fromEmailPassword);
            }
         });

      try {
         // Create a default MimeMessage object.
         Message message = new MimeMessage(session);

         // Set From: header field of the header.
         message.setFrom(new InternetAddress(AppConfig.fromEmail));

         // Set To: header field of the header.
         message.setRecipients(Message.RecipientType.TO,
            InternetAddress.parse("gopi.v@beetlerim.com"));

         // Set CC: header field of the header.
         message.setRecipients(Message.RecipientType.CC,
            InternetAddress.parse("mahitha.r@beetlerim.com"));
         
         
         // Set Subject: header field
         String subject = "Daily Marketing Report by Agent - "+agentName;
         message.setSubject(subject);

         // Create the message part
         BodyPart messageBodyPart = new MimeBodyPart();

         // Now set the actual message
         String messageInBody = "Hi "+agentName+", \n\n Please find the attachment.\n\n Regards,\n HomeRiver Group.";
         messageBodyPart.setText(messageInBody);

         // Create a multipar message
         Multipart multipart = new MimeMultipart();

         // Set text message part
         multipart.addBodyPart(messageBodyPart);

         // Part two is attachment
         messageBodyPart = new MimeBodyPart();
         String filename = AppConfig.downloadPath+"_"+agentName+".xlsx";
         System.out.println("FileName sending in mail"+filename);
         messageBodyPart.setFileName(new File(filename).getName());
         DataSource source = new FileDataSource(filename);
         messageBodyPart.setDataHandler(new DataHandler(source));
        // messageBodyPart.setFileName(filename);
         messageBodyPart.setFileName(new File(filename).getName());
         multipart.addBodyPart(messageBodyPart);

         // Send the complete message parts
         message.setContent(multipart);

         // Send message
         Transport.send(message);

         System.out.println("Sent message successfully....");
  
         //wait until file is downloaded
         /*
         File dir = new File("DownloadPath");
         //String partialName = downloaded_report.split("_")[0].concat("_"); //get cancelled and add underscore
        // FluentWait<WebDriver> wait = new FluentWait<WebDriver>(RunnerClass.driver);
                 //wait.pollingEvery(1, TimeUnit.SECONDS);
                 //wait.withTimeout(15, TimeUnit.SECONDS);
                 RunnerClass.wait.until(x -> {
                     File[] filesInDir = dir.listFiles();
                     for (File fileInDir : filesInDir) {
                         if (fileInDir.getName().startsWith("Marketing")) {
                             return true;
                         }
                     }
                     return false;
                 });
         */
         //delete the current file
         File file = new File(filename);
         file.delete();
      } catch (MessagingException e) 
      {
    	  e.printStackTrace();
         throw new RuntimeException(e);
      }
   }
}
