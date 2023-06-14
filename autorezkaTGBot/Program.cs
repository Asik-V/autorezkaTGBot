using System.IO;
using System.Net.WebSockets;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using System.ComponentModel;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Diagnostics;
using System.Net.Http;

namespace autorezkaTGBot
{
    class Program
    {
        static ITelegramBotClient bot = new TelegramBotClient("6210634318:AAEfSC2S4Sv9Zh3ndWn_w-p73zpwTaKhiCU");
        static void Main(string[] args)
        {
            Console.WriteLine($"Bot - {bot.GetMeAsync().Result.Username}");

            var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { },
            };
            bot.StartReceiving(
                HandleUpdateAsync,
                HandleErrorAsync,
                receiverOptions,
                cancellationToken
            );
            Console.ReadLine();


        }

        public static void Logging(Message message, string nameFolder, List<long> copyArr, string status)
        {
            System.IO.File.AppendAllText("log.txt", $"\n\n\nUsername: {message.From.Username}\n" +
                                                                $"ChatID: {message.Chat.Id}\n" +
                                                                $"Name Folder: {nameFolder}\n" +
                                                                $"Phone Numbers: {copyArr.Count}\n" +
                                                                $"Time: {message.Date.ToString("G")}\n" +
                                                                $"Status: {status}");
        }

        public static void Logging(Message message, string status)
        {
            System.IO.File.AppendAllText("log.txt", $"\n\n\nUsername: {message.From.Username}\n" +
                                                                $"ChatID: {message.Chat.Id}\n" +
                                                                $"Message: {message.Text}\n" +
                                                                $"Time: {message.Date.ToString("G")}\n" +
                                                                $"Status: {status}");
        }

        public static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            
            
            
            if (update.Type == Telegram.Bot.Types.Enums.UpdateType.Message)
            {


                var message = update.Message;
                Console.WriteLine($"\n\n\nChatID: {message.Chat.Id}");
                Console.WriteLine($"Username: {message.From.Username}");

                //заполнение текстового документа диапазонами из бота
                System.IO.File.WriteAllText(@"..\..\..\text.txt", message.Text);

                //заполнение массива из текстового документа
                string[] arrStr = System.IO.File.ReadAllText(@$"C:\Users\{Environment.UserName}\Desktop\autorezkaTGBot\autorezkaTGBot\text.txt").Split('\n');
                
                string nameFolder = arrStr[arrStr.Length - 1];
                nameFolder = nameFolder.Replace(' ', '_');
                Console.WriteLine($"Название папки, екселей: {nameFolder}");

                
                string[] indexOne = arrStr[0].Split(' ');
                if (indexOne[0].Length == 11)
                {

                    Directory.CreateDirectory(@$"..\..\..\{nameFolder}");
                    Array.Resize<string>(ref arrStr, arrStr.Length - 1);

                    List<long> completeArr = new List<long>();
                    List<long> copyArr = new List<long>();
                    Random random = new Random();
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;



                    //алгоритм растяжки
                    Console.WriteLine("растяжка");
                    for (int i = 0; i <= arrStr.Length - 1; i++)
                    {
                        string[] numbers = arrStr[i].Split(Char.Parse(" "));
                        for (long j = long.Parse(numbers[0]); j <= long.Parse(numbers[2]); j++)
                        {
                            completeArr.Add(j);
                        }
                        System.Console.WriteLine($"{numbers[0]} - {numbers[2]}");
                    }
                    copyArr.AddRange(completeArr);
                    //перемешивание
                    Console.WriteLine("перемешивание");
                    for (int i = completeArr.Count - 1; i >= 1; i--)
                    {
                        int j = random.Next(i + 1);
                        var temp = completeArr[j];
                        completeArr[j] = completeArr[i];
                        completeArr[i] = temp;
                    }

                    int countExcels = completeArr.Count / 50000;


                    //заполнение екселей данными массива
                    Console.WriteLine("заполнение");
                    for (int i = 0; i <= countExcels - 1; i++)
                    {
                        using (ExcelPackage excelPackage = new ExcelPackage())
                        {
                            excelPackage.Workbook.Worksheets.Add("List1");
                            excelPackage.Workbook.Worksheets[0].Cells["A1"].Value = 1;
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];

                            for (int j = 0; j < 50000; j++)
                            {
                                worksheet.Cells[$"A{j + 1}"].Value = completeArr[j];
                            }

                            completeArr.RemoveRange(0, 50000);

                            excelPackage.SaveAs(@$"C:\\Users\\{Environment.UserName}\\Desktop\\autorezkaTGBot\\autorezkaTGBot\\{nameFolder}\\{nameFolder}-{i + 1}.xlsx");

                        }
                    }


                    //архивация папки с екселями
                    Console.WriteLine("архивация");
                    string arg = $@"a ..\..\..\{nameFolder}.rar ..\..\..\{nameFolder}";
                    ProcessStartInfo ps = new ProcessStartInfo();
                    ps.FileName = @"C:\Program Files\WinRAR\WinRAR.exe";
                    ps.Arguments = arg;
                    Process.Start(ps);
                    Thread.Sleep(3000);

                    //удаление папки с екселями
                    Directory.Delete($@"..\..\..\{nameFolder}", true);

                    //отправка архива на сторону клиента
                    Console.WriteLine("отправка");
                    using (Stream stream = System.IO.File.OpenRead($@"..\..\..\{nameFolder}.rar"))
                    {
                        await botClient.SendDocumentAsync(message.Chat.Id, new InputFileStream(content: stream, fileName: $@"..\..\..\{nameFolder}.rar"));
                    }
                    Thread.Sleep(5000);
                    System.IO.File.Delete(@$"..\..\..\{nameFolder}.rar");

                    Console.WriteLine("Готово");
                    
                    Logging(message, nameFolder, copyArr, "complete");
                    
                    return;
                }
                else if (message.Text.ToLower() == "/start")
                {
                    await botClient.SendTextMessageAsync(message.Chat, "Кидай заготовки, Пых");
                    Logging(message, "/start");
                    return;
                }
                else
                {
                    await botClient.SendTextMessageAsync(message.Chat, "Неверно указан формат");
                    Logging(message, "error");
                    return;
                }
            }
        }

        public static async Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            // Некоторые действия
            Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(exception));
        }
    }
}
