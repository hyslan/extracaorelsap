using System;
using System.IO;

namespace Testes
{
    class Program
    {
        static void Main(string[] args)
        {
            string pastaDeTrabalho = Directory.GetCurrentDirectory();
            Console.WriteLine("A pasta de trabalho do projeto é: " + pastaDeTrabalho);
        }
    }
}
