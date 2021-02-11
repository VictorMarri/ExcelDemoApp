using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            //EPPlus é pago em contexos comerciais, aqui estamos falando que seu uso não será para fins comerciais
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Esse é o caminho do arquivo, onde vai ficar alocado os arquivos que criarmos
            var file = new FileInfo(@"C:\Demos\ExcelDemo.xlsx");

            //Aqui a gente ja vai estar dando um 'Load' na nossa lista de Pessoas
            var people = GetSetupData();

            //Vamos salvar os dados no arquivo, e como isso pode demorar um certo tempo, usamos o await e transformamos o Main em async Task
            await SaveExcelFile(people, file); 
        }

        /// <summary>
        /// Metodo que vai realizar o salvamento do arquivo, com as informações dentro
        /// </summary>
        /// <param name="people">A Lista de dados que criamos para ser inserida dentro do excel</param>
        /// <param name="file">O Arquivo em si, o seu caminho. Será feita checagem pra ver se ja existe</param>
        /// <returns>Retorna o Salvamento do Arquivo</returns>
        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            //Checa se o arquivo q foi passado existe. Se existe, vai deletar esse arquivo antes da gente rodar a aplicação
            DeleteIfExists(file);

            //Qualquer coisa que instanciarmos aqui, no final da sua execução, vai ter seus recursos limpos. Sem registros dele soltos pela aplicação, caso alguem abra fisicamente o arquivo e ele esteja impedido pq algo na aplicação deixou resquicios. Ou seja, using vai chamar automaticamente o metodo Disposable() pra gente no final desse escopo.
            using var package = new ExcelPackage(file);

            //To colocando um Sheet chamado "MainReport" no excel file
            var ws = package.Workbook.Worksheets.Add("MainReport");

            //Entao agora a gente tem um sheet novo no nosso excel chamado "MainReport", agora a gente vai adicionar dados nesse sheet
            var range = ws.Cells["A1"].LoadFromCollection(people, PrintHeaders: true); //Vai pegar os dados da nossa lista de pessoas que criamos, e vai salvar no excel como colunas.E a primeira linha vai ser a linha de cabeçalho, que pega o nome da propriedade e coloca na primeira coluna  

            //To ajustando as colunas, os tamanhos dela de acordo com o tamanho dos dados q tiver la 
            range.AutoFitColumns(); 

            await package.SaveAsync();
        
        }

        /// <summary>
        /// Metodo que realiza a exclusão do arquivo, caso ele ja exista quando formos executar a aplicação.
        /// O arquivo vai ser excluido só pra fins de teste mesmo. De jeito algum fariamos isso em uma aplicação de valor.
        /// </summary>
        /// <param name="file">O Arquivo para realizarmos a checagem caso exista</param>
        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        /// <summary>
        /// Esses sao os dados que vamos estar inserindo no nosso excel, para fins de demonstração
        /// </summary>
        /// <returns>Retorna uma lista de dados, dados estes que vao para dentro do arquivo excel</returns>
        private static List<PersonModel> GetSetupData() //Estatico pq vamos usar esse metodo dentro do main, e o main é estatico, e estatico só aceita estatico
        {
            List<PersonModel> output = new() //Esse New() é só uma forma de simplificar a instancia da classe. Daria na mesma eu usar = New List<PersonModel>(), mas assim fica muito mais simples
            {
                new() { Id = 1, FirstName = "Victor", LastName = "Marri"},
                new() { Id = 2, FirstName = "Bruce", LastName = "Lee" },
                new() { Id = 3, FirstName = "Brandon", LastName = "Lee" },
                new() { Id = 4, FirstName = "Leonardo", LastName = "Da Vinci" },
                new() { Id = 5, FirstName = "Leonardo", LastName = "Di Caprio" },
                new() { Id = 6, FirstName = "Rosa", LastName = "Diaz" },
                new() { Id = 7, FirstName = "Carl", LastName = "Johnson" }

            };

            return output;

        }
    }
}
