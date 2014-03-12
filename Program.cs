using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;//Esta es necesaria
using System.Text.RegularExpressions; //Esta es necesaria en visual solo se cambia el using por el import
using BCL.easyPDF7.Interop.EasyPDFDocument;
using BCL.easyPDF7.Interop.EasyPDFLoader;
using BCL.easyPDF7.Interop.EasyPDFProcessor;

//Imports System.Text.RegularExpressions
namespace PdfPages
{    
    class Program
    {
        static string archivo = "C:\\Users\\jhonatan.bazalduao\\Documents\\Visual Studio 2010\\Projects\\PdfPages\\PdfPages\\PDF\\Python.pdf";         
        /*Retorna un entero con el numero de paginas 
            lo que hace es crear un nuevo objeto StreamReader, con el regex crea la expresion regular y con match busca las coincidencias
            si no mal recuerdo y tengo mucho que no uso Vb se declara así 
            "Using Stream as StreamReader = New StreamReader("file.pdf")"
            y lo de mas es lo mismo.
        */
        static public int getNumeroDePaginas(string fileName)
        {           
            using (StreamReader sr = new StreamReader(File.OpenRead(archivo)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());
                return matches.Count;
            }
        }
        static void Main(string[] args)
        {            
            Console.WriteLine("son {0} paginas ",getNumeroDePaginas(archivo));
            Console.ReadLine();
        }        
    }
}
