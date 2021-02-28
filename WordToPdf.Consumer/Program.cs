using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace WordToPdf.Consumer
{
    class Program
    {

        public static bool EmailSend(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0;

                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attachment = new Attachment(memoryStream, ct);

                attachment.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage();

                SmtpClient smtpClient = new SmtpClient();

                mailMessage.From = new MailAddress("s.oykubilen@gmail.com");

                mailMessage.To.Add(email);

                mailMessage.Subject = "Pdf Dosyası";

                mailMessage.Body = "pdf dosyasınız ektedir";
                mailMessage.IsBodyHtml = true;
                mailMessage.Attachments.Add(attachment);
                smtpClient.Port = 587;
                smtpClient.Host = "smtp.gmail.com";
                smtpClient.UseDefaultCredentials = true;
                smtpClient.Credentials = new System.Net.NetworkCredential("", "");//mail and password
                smtpClient.EnableSsl = true;
                smtpClient.Send(mailMessage);
                Console.WriteLine($"Sonuc:{email} adresine gönderilmiştir.");
                Console.ReadLine();
                memoryStream.Close();
                memoryStream.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Sonuc:{email} adresine gönderilme sırasında hata meydana geldi.Hata:{ex.InnerException}");
                Console.ReadLine();

                return false;
            }

        }
        static void Main(string[] args)
        {
            bool result = false;
            var factory = new ConnectionFactory { HostName = "localhost" };
            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);

                    channel.QueueBind("File", "convert-exchange", "WordToPdf", null);//exchange ile que bağladık

                    channel.BasicQos(0, 1, false);//mesajların eşit şekilde alınmasını sağlar,aynı anda 2 mesaj gelmesini engelledik

                    var consumer = new EventingBasicConsumer(channel);

                    //autoack: işlem yapılamadığında mesaj düşssün mü? false=düşmesin, kendim belirticem düşüp düşmediğini
                    channel.BasicConsume("File", false, consumer);

                    consumer.Received += (model, ea) =>
                     {
                         try
                         {
                             Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor");

                             Document doc = new Document();
                             string message = Encoding.UTF8.GetString(ea.Body.ToArray());
                             MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(message);

                             doc.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Odt);

                             using (MemoryStream ms = new MemoryStream())
                             {
                                 doc.SaveToStream(ms, FileFormat.PDF);
                                 result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);
                             }
                         }
                         catch (Exception ex)
                         {
                             Console.WriteLine($"Hata meydana geldi. {ex.Message}");
                             result = false;
                             throw;
                         }

                         if (result)
                         {
                             Console.WriteLine("Kuyruktan  mesaj başarıyla işlendi");
                             channel.BasicAck(ea.DeliveryTag, false);//başarılıysa sadece bu mesaj gitsin ve kuyruktan silinsin
                         }

                     };

                    Console.WriteLine("Çıkmak için tıklayınız");
                    Console.ReadLine();
                }
            }
        }
    }
}
