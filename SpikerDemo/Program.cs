using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace SpikerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
           string text = File.ReadAllText("DonQuijote.txt");
           if(!string.IsNullOrEmpty(text))            
              text = Speech.cleantext(text);

            Speech.Speak(text);
        }
    }

    public static class Speech
    {
        public static string cleantext(string text)
        {
            //obtencion de repeticiones
            var arr = text.Split('\n');

            if(arr.Length <= 1) return arr[0];
            
            StringBuilder builder = new StringBuilder();

            foreach (var item in arr)
            {
                if (item.Contains(" x"))
                {
                    //texto
                    string partetexto = item.Split('x')[0];
                    //cantidad
                    string partecantidad = item.Split('x')[1];

                    if (partecantidad.Length > 2)
                    {
                        partetexto = partetexto + partecantidad.Substring(2);
                        partecantidad = partecantidad.Substring(0, 2);
                    }

                    if (int.TryParse(partecantidad, out int cantidad))
                    {
                        for (int i = 0; i < cantidad; i++)
                        {
                            builder.Append(partetexto + "\n");
                        }
                    }
                }
                else
                {
                    builder.Append(item + "\n");
                }
            }


            return builder.ToString();
        }

        public static void Speak(string text, bool wait = false)
        {
            //Power shell para convertir texto a audio...
            ExecuteCommand(
                $@"Add-Type -AssemblyName System.speech; 
                $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer; 
                $speak.Speak(""{text}"");");

            void ExecuteCommand(string command)
            {
                string path = Path.GetTempPath() + Guid.NewGuid() + ".ps1";

                // make sure to be using System.Text
                using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
                {
                    sw.Write(command);

                    ProcessStartInfo start = new ProcessStartInfo()
                    {
                        FileName = @"C:\Windows\System32\windowspowershell\v1.0\powershell.exe",
                        LoadUserProfile = false,
                        UseShellExecute = false,
                        CreateNoWindow = true,                        
                        Arguments = $"-executionpolicy bypass -File {path}",
                        WindowStyle = ProcessWindowStyle.Hidden
                    };

                    Process process = Process.Start(start);

                   if (wait)
                       process.WaitForExit();
                }
            }
        }
    }
}