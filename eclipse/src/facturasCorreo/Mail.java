package facturasCorreo;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class Mail {
	private String correoDeOrigen;
    private String correoDeDestino;
    private String asunto;
    private String mensajeDeTexto;
    private String password;
    private int year;
            
    public Mail(String origen,String destino,String asunto, String txt,String password, int year){
    	this.correoDeOrigen = origen;
    	this.correoDeDestino = destino;
    	this.asunto = asunto;
    	this.mensajeDeTexto = txt;
    	this.password = password;   
        this.year = year;
    }
    
    
    public static void main(String[] args) {
    	String to = "pablogl2002@gmail.com";
    	System.out.println("Preparando...");
    	envioDeCorreos(to, "marzo", 7, "black ice.xlsx", 2022);
    	System.out.println("Enviado");
    }
    
    public static void envioDeCorreos(String to,String month, int nFact, String file, int year){
		String from = "muebledos@outlook.com";
		String pwd = "mueble1963";
		
		String subject = "Factura " + month;
		String txt = "Adjunto factura nº " + nFact + " correspondiente al mes de " + month + " de "+ year + "\n" + "\n"
              + "Un saludo, Enrique García.\n" ;
		
		Mail m = new Mail(from, to, subject, txt, pwd, year);
	  	m.envioDeMensajes(file);
    }
    
            
    private void envioDeMensajes(String file){
    	try{
    		//System.out.println("Antes propiedades");
    		Properties props = new Properties();  
    		props.put("mail.smtp.socketFactory.fallback", "false");  
    		props.put("mail.smtp.quitwait", "false");
    		props.put("mail.smtp.socketFactory.port", "587");  
    		props.put("mail.host", "smtp.office365.com");
			     
    		props.setProperty("mail.transport.protocol", "smtp");
				     
    		props.setProperty("mail.smtp.port", "587");
    		props.setProperty("mail.smtp.ssl.trust", "*");
    		props.setProperty("mail.smtp.starttls.enable", String.valueOf(true));//True or False
    		props.setProperty("mail.smtp.ssl.protocols", "TLSv1.2");
    		props.setProperty("mail.smtp.timeout", "300000");
    		props.setProperty("mail.smtp.connectiontimeout", "300000");
    		props.setProperty("mail.smtp.writetimeout", "300000");
			  
    		Session s = Session.getDefaultInstance(props);
    		
    		//System.out.println("Despues propiedades, antes mensaje");
    		
    		//adjunto
    		BodyPart texto = new MimeBodyPart();
    		texto.setText(mensajeDeTexto);
    		BodyPart adjunto = new MimeBodyPart();
    		//direccion USB
    		//adjunto.setDataHandler(new DataHandler(new FileDataSource("D:/CLIENTES/año "+ year + "/pdf/" + file)));
    		//direccion carlos
    		adjunto.setDataHandler(new DataHandler(new FileDataSource("C:/Users/carlo/Desktop/CLIENTES/año "+ year + "/pdf/" + file)));
    		//direccion papa
    		//adjunto.setDataHandler(new DataHandler(new FileDataSource("C:/Users/pablo/Desktop/CLIENTES/año "+ year + "/pdf/" + file)));
    		adjunto.setFileName(file);
    		MimeMultipart m = new MimeMultipart();
    		m.addBodyPart(texto);
    		m.addBodyPart(adjunto);
    		 
    		
    		//Mensaje texto
    		MimeMessage mensaje = new MimeMessage(s);
    		mensaje.setFrom(new InternetAddress(correoDeOrigen));
    		mensaje.addRecipient(Message.RecipientType.TO, new InternetAddress(correoDeDestino));
    		mensaje.setSubject(asunto);
    		//mensaje.setText(mensajeDeTexto);
    		mensaje.setContent(m);
		  
    		//System.out.println("Despues mensaje, antes envio");
		  
    		Transport t = s.getTransport("smtp");
    		t.connect(correoDeOrigen,password);
    		t.sendMessage(mensaje, mensaje.getAllRecipients());
    		t.close();
    		//System.out.println("exito");
    	} catch (MessagingException e) {
    	  	System.out.print(e);
    	}
    }
}
