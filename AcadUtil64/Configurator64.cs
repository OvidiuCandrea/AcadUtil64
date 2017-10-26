using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;

namespace Ovidiu.x64.General
{
    public class Configurator
    {

        //Definire Field-uri
        string caleFisierConfig = null;
        List<string> listaRez = new List<string>();

        //Definire Proprietati
        public bool FisierulExista
        {
            get
            {
                return File.Exists(caleFisierConfig);
            }
        }

        //Constructor - seteaza calea fisierului de configurare
        public Configurator(string cale)
        {
            //Setarea caii fisierului de configurare, dupa verificarea/corectarea formatului acesteia
            caleFisierConfig = Path.GetFullPath(cale);
        }

        public string[] citireConf(string cod)
        {
            string[] rezultat;

            try
            {
                using (StreamReader cititor = new StreamReader(caleFisierConfig))
                {
                    List<string> continut = new List<string>();

                    string linie;
                    int contor = 0;

                    while ((linie = cititor.ReadLine()) != null)
                    {
                        continut.Add(linie);
                        contor++;
                    }

                    cititor.Dispose();

                    contor = -1; //Resetare contor
                    bool scrie = false; //Se scrie sau nu in matricea rezultat//

                    foreach (string rand in continut)
                    {
                        if (rand.StartsWith(cod))
                        {
                            if (scrie) scrie = false;
                            else scrie = true;
                        }
                        else if (scrie && rand.StartsWith("#") == false)
                        {
                            //Adauga randul la lista Rezultat, fara un potential comentariu
                            if (rand.Contains("#")) listaRez.Add(rand.Remove(rand.IndexOf("#")));
                            else listaRez.Add(rand);
                        }
                    }
                }
            }
            catch
            {
            }

            rezultat = listaRez.ToArray();

            return rezultat;
        }

        public void scriereConf(string cod, string[] optiuni)
        {
            //Citirea optiunilor existente, cu codul dat, in fisierul de configurare
            string[] optExistente = citireConf(cod);

            //Impartirea optiunilor in optiuni existente si optiuni noi
            bool existenta = false;
            List<string> optNoi = new List<string>();
            Hashtable perechiOpt = new Hashtable();
            foreach (string optiune in optiuni)
            {
                existenta = false;
                if (optiune.Contains("="))
                {
                    foreach (string optExistenta in optExistente)
                    {
                        if (optExistenta.StartsWith(optiune.Substring(0, optiune.IndexOf("="))))
                        {
                            perechiOpt[optExistenta] = optiune;
                            existenta = true;
                        }
                    }
                }
                if (existenta == false)
                {
                    optNoi.Add(optiune);
                }
            }

            //Scrierea optiunilor
            string deScris = "";
            using (StreamReader cititor = new StreamReader(caleFisierConfig))
            {
                deScris = cititor.ReadToEnd();
            }
            using (StreamWriter scriitor = new StreamWriter(caleFisierConfig))
            {

                //Adaugarea codului daca acesta nu este gasit
                if (deScris.Contains(cod) == false)
                {
                    deScris = deScris + "\r\n" + cod + "\r\n" + cod + "\r\n";
                }

                //Inlocuirea optiunilor existente cu unele noi
                if (perechiOpt.Count > 0)
                {
                    foreach (DictionaryEntry pereche in perechiOpt)
                    {
                        deScris = deScris.Replace((string)pereche.Key, (string)pereche.Value);
                    }
                }

                //Adaugarea optiunilor complet noi
                if (optNoi.Count > 0)
                {
                    foreach (string optiune in optNoi)
                    {
                        deScris = deScris.Insert(deScris.LastIndexOf(cod) - 1, "\n" + optiune + "\r\n");
                    }
                }

                scriitor.Write(deScris);
                scriitor.Dispose();
            }

        }

    }
}
