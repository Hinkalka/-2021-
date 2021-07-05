using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Экзамен_2021_ПП
{
    class Program
    {
        static void Main(string[] args)
        {
        izm7: string listname = string.Empty;
            Console.WriteLine("Какой лист хотите выбрать? ");
            switch (Console.ReadLine())
            {
                case "1":
                    Console.WriteLine("Выбран 1 лист. Подождите минуту.");
                    listname = "Лист1";
                    break;
                case "2":
                    Console.WriteLine("Выбран 2 лист. Подождите минуту.");
                    listname = "Лист2";
                    break;
                case "3":
                    Console.WriteLine("Выбран 3 лист. Подождите минуту.");
                    listname = "Лист3";
                    break;
                case "4":
                    Console.WriteLine("Выбран 4 лист. Подождите минуту.");
                    listname = "Лист4";
                    break;
                default:
                    Console.WriteLine("Такого листа нет.");
                    break;
            }
        }
        string filename = @"C:\Users\1\Downloads\komivoyazher.xlsx";
        int column = 0;
        double[,] table = Excel.GetArray(filename, listname, out column);
        double sum = 0;

    }
    public class konec
    {
        public void k()
        {
            var g = new g();

            //добавление вершин
            g.AddVertex("1");
            g.AddVertex("2");
            g.AddVertex("3");
            g.AddVertex("4");
            g.AddVertex("5");


            //добавление ребер
            g.AddEdge("1", "2", 3);
            g.AddEdge("1", "3", 7);
            g.AddEdge("1", "4", 1);
            g.AddEdge("2", "4", 8);
            g.AddEdge("3", "4", 9);
            g.AddEdge("4", "5", 12);

            Console.WriteLine(g);
            var dijkstra = new D(g);
            var path = dijkstra.FindShortestPath("1", "5");
            Console.WriteLine(path);

        }
    }

}
