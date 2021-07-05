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
    public class gvi1
    {
        public gv v { get; set; }
        public bool IU { get; set; }
        public int ews { get; set; }
        public gv pv { get; set; }
        public gvi1(gv v)
        {
            this.v = v;
            IU = true;
            ews = int.MinValue;
            pv = null;
        }
    }
    public class g
    {
        public List<gv> Vertices { get; }
        public g()
        {
            Vertices = new List<gv>();
        }
        public void AddVertex(string vertexName)
        {
            Vertices.Add(new gv(vertexName));
        }
        public gv FindVertex(string vertexName)
        {
            foreach (var v in Vertices)
            {
                if (v.Name.Equals(vertexName))
                {
                    return v;
                }
            }

            return null;
        }
        public void AddEdge(string firstName, string secondName, int weight)
        {
            var v1 = FindVertex(firstName);
            var v2 = FindVertex(secondName);
            if (v2 != null && v1 != null)
            {
                v1.AddEdge(v2, weight);
                v2.AddEdge(v1, weight);
            }
            Console.WriteLine("{0},{1}", v1, v2);
        }
    }

    public class D
    {
        g g;
        List<gvi1> infos;
        public D(g graph)
        {
            this.g = graph;
        }
        void InitInfo()
        {
            infos = new List<gvi1>();
            foreach (var v in g.Vertices)
            {
                infos.Add(new gvi1(v));
            }
        }
        gvi1 GetVertexInfo(gv v)
        {
            foreach (var i in infos)
            {
                if (i.v.Equals(v))
                {
                    return i;
                }
            }

            return null;
        }
        public gvi1 FindUnvisitedVertexWithMinSum()
        {
            var maxValue = int.MinValue;
            gvi1 maxVertexInfo = null;
            foreach (var i in infos)
            {
                if (i.IU && i.ews > maxValue)
                {
                    maxVertexInfo = i;
                    maxValue = i.ews;
                }
            }
            return maxVertexInfo;
        }
        public string FindShortestPath(string startName, string finishName)
        {
            Console.WriteLine("Начало пути: {0}", startName);
            Console.WriteLine("Конец пути: {0}", finishName);
            return FindShortestPath(g.FindVertex(startName), g.FindVertex(finishName));
        }
        public string FindShortestPath(gv startVertex, gv finishVertex)
        {
            InitInfo();
            var first = GetVertexInfo(startVertex);
            first.ews = 0;
            while (true)
            {
                var current = FindUnvisitedVertexWithMinSum();
                if (current == null)
                {
                    break;
                }

                SetSumToNextVertex(current);

            }

            return GetPath(startVertex, finishVertex);
        }
        void SetSumToNextVertex(gvi1 info)
        {
            info.IU = false;
            foreach (var e in info.v.Edges)
            {
                var nextInfo = GetVertexInfo(e.ConnectedVertex);
                var sum = info.ews + e.EdgeWeight;
                if (sum > nextInfo.ews)
                {
                    nextInfo.ews = sum;
                    nextInfo.pv = info.v;
                    Console.WriteLine("Сумма после добавления предыдущего временного промежутка: {0}", sum);
                }
            }

        }
        public string GetPath(gv startVertex, gv endVertex)
        {
            var path = endVertex.ToString();
            while (endVertex == null)
            {
                endVertex = GetVertexInfo(endVertex).pv;
                path = "Номер вершины:" + endVertex.ToString() + "; " + path;
            }
            return path;
        }
        void fl(string p)
        {

            StreamWriter sw;

        }
        void wtf(string p1)
        {
            using (StreamWriter sw = new StreamWriter(p1, false))
            {
                StreamReader sr;
                const int NmaxZap = 10;
                sr = new StreamReader(@"\\main\RDP\31П\СергеевДИ\Desktop\практика.txt", UTF8Encoding.Default);
                string[] d = new string[NmaxZap];
                string t = sr.ReadLine();
                int i = 0;
                while ((t != null) && (i < d.Length))
                {
                    Console.WriteLine(t);
                    d[i++] = t;
                    t = sr.ReadLine();
                }
                sr.Close();
            }
        }
    }

}



