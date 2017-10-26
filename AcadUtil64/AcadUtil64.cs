using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows;
//using Autodesk.Civil.Runtime;
//using Autodesk.Civil.ApplicationServices;
//using Autodesk.Civil.DatabaseServices;
//using Autodesk.Civil.Settings;
using Ovidiu.StringUtil;

namespace Ovidiu.x64.AcadUtil
{
    public class ModUtil
    {

    }

    public class FileUtil
    {
        //Comanda pentru interpolare liniara a unui fisier STO
        [CommandMethod("LINTERP")]
        public void linterp()
        {
            Document acadDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acadDoc.Editor;
            //Se creaza lista de puncte noi, ale caro cote vor fi calculate prin interpolare
            String3D listanoua = new String3D();

            //Descriere
            ed.WriteMessage("\nCommand for linear interpolation on station-offset files");

            //Permiterea sau interzicerea extrapolarilor
            PromptKeywordOptions PKO2 = new PromptKeywordOptions("\nAllow extrapolations?");
            PKO2.Keywords.Add("Yes");
            PKO2.Keywords.Add("No");
            PKO2.Keywords.Default = "Yes";
            PKO2.AppendKeywordsToMessage = true;
            PromptResult PKR2 = ed.GetKeywords(PKO2);

            //Descrierea punctelor noi
            PromptKeywordOptions PKO3 = new PromptKeywordOptions("\nNew points description");
            PKO3.Keywords.Add("linterp");
            PKO3.Keywords.Default = "linterp";
            PKO3.AppendKeywordsToMessage = true;
            PKO3.AllowArbitraryInput = true;
            string descr = ed.GetKeywords(PKO3).StringResult;

            //selectia fisierului STO cu date
            ed.WriteMessage("\nData source file selection:");
            PromptOpenFileOptions PFO = new PromptOpenFileOptions("\nSelect station-offset file: ");
            PFO.Filter = "Text file (*.txt)|*.txt|Comma separated file (*.csv)|*.csv";
            PromptFileNameResult PFR = ed.GetFileNameForOpen(PFO);
            if (PFR.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nInvalid file selection! Aborting.");
                return;
            }
            string startfile = PFR.StringResult;

            //Selectarea operatiunii de interpolare dorita
            PromptKeywordOptions PKO = new PromptKeywordOptions("");
            PKO.Message = "\nResult file to contain all points or only newly interpolated ones?";
            PKO.Keywords.Add("All");
            PKO.Keywords.Add("New");
            PKO.AppendKeywordsToMessage = true;
            PromptResult PKR = ed.GetKeywords(PKO);
            if (PKR.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nAborting.");
                return;
            }
                        
            //Citirea tuturor punctelor din fisierul sursa si sortarea lor dupa km si offset
            String3D listapuncte = new String3D();
            listapuncte.ImportPoints(startfile, Punct3D.Format.KmOZD, Punct3D.DelimitedBy.Comma);
            if (listapuncte.Count < 2)
            {
                ed.WriteMessage("\nNot enough points were read from the selected file! Aborting.");
                return;
            }
            listapuncte.Sort(
                delegate(Punct3D p1, Punct3D p2)
                {
                    if (p1.KM.CompareTo(p2.KM) == 0) return p1.Offset.CompareTo(p2.Offset);
                    else return p1.KM.CompareTo(p2.KM);
                }
            );

            //Impartirea punctelor pe sectiuni
            List<String3D> sectiuni = new List<String3D>();
            String3D sectiune = new String3D();
            sectiune.Add(listapuncte[0]);
            for (int i = 1; i < listapuncte.Count; i++)
            {
                if (listapuncte[i].KM != listapuncte[i - 1].KM)
                {
                    sectiuni.Add(sectiune);
                    sectiune = new String3D();
                }
                sectiune.Add(listapuncte[i]);
            }
            sectiuni.Add(sectiune);
            ed.WriteMessage("\n{0} points were read from the source file, arranged in {1} sections from km {2} to km {3}",
                listapuncte.Count, sectiuni.Count, sectiuni[0][0].KM, sectiuni[sectiuni.Count - 1][0].KM);


            //Citirea offseturilor introduse de la tastatura sau alegerea metodei fisierului
            //Realizarea listei de puncte noi
            ed.WriteMessage("\nOffsets can be specified by value or by difference from  first or last section point (e.g. min-3.25;max+0.7)");
            PromptKeywordOptions PSO = new PromptKeywordOptions("\nSpecify the desired offsets to be added, separated by semicolon");
            PSO.Keywords.Add("fromFile");
            PSO.Keywords.Add("eXit");
            PSO.AllowArbitraryInput = true;
            PSO.AllowNone = false;
            PSO.AppendKeywordsToMessage = true;
            PromptResult PSR = ed.GetKeywords(PSO);

            

            //Operatiune nepermisa sau parasirea programului
            if (PSR.Status == PromptStatus.Error || PSR.Status == PromptStatus.Cancel || PSR.StringResult == "eXit")
            {
                ed.WriteMessage("\nAborting.");
                return;
            }

            //Metoda fisierului STO
            if (PSR.StringResult == "fromFile")
            {
                PFO.DialogName = "Station-offset file selection";
                PFO.DialogCaption = "Select station-offset file";
                PromptFileNameResult PFR2 = ed.GetFileNameForOpen(PFO);
                if (PFR2.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nInvalid file selection! Aborting.");
                    return;
                }
                String3D punctetinta = new String3D();
                punctetinta.ImportPoints(PFR2.StringResult, Punct3D.Format.KmOZD, Punct3D.DelimitedBy.Comma);
                foreach (Punct3D P in punctetinta)
                {
                    try
                    {
                        if (listapuncte.Find(punct => punct.KM == P.KM) != null)
                        {
                            P.Z = -999;
                            P.D = descr;
                            listanoua.Add(P);
                        }
                    }
                    catch { }
                }
                ed.WriteMessage("\n{0} points read from file", listanoua.Count);
            }

            //Metoda introducerii de la tastatura
            else if (PSR.Status == PromptStatus.OK)
            {
                string[] distStr;
                if (PSR.StringResult.Contains(';')) distStr = PSR.StringResult.Split(';');
                else distStr = new string[] { PSR.StringResult };
                //ed.WriteMessage("\n{0} offsets will be added to each section", distStr.Length);

                //Parcurgerea sectiunilor transversale
                foreach (String3D sect in sectiuni)
                {
                    ////ed.WriteMessage("\nProcessing km {0}", sect[0].KM);
                    //if (sect.Count < 2 || sect[sect.Count - 1].Offset - sect[0].Offset == 0)
                    //{
                    //    ed.WriteMessage("\nNot enough points for interpolation at km {0}! Section skipped.", sect[0].KM);
                    //    continue;
                    //}

                    //Parcurgerea distantelor introduse de la tastatura si completarea listei noi de puncte
                    foreach (string distanta in distStr)
                    {
                        //ed.WriteMessage("\nParsing offset " + distanta);
                        //Obtinerea distantei curente ca "double"
                        double offset;
                        if (distanta.Contains("min"))
                        {
                            if (double.TryParse(distanta.Replace("min", ""), out offset) != false)
                            {
                                offset = sect[0].Offset + offset;
                                Punct3D p = new Punct3D();
                                p.KM = sect[0].KM;
                                p.Offset = offset;
                                p.Z = -999;
                                p.D = descr;
                                //ed.WriteMessage(" ---->");
                                listanoua.Add(p);
                                //ed.WriteMessage(" " + offset.ToString());
                            }
                            else
                            {
                                ed.WriteMessage("Offset {0} is invalid and will be ignored.", distanta);
                                continue;
                            }
                        }
                        else if (distanta.Contains("max"))
                        {
                            if (double.TryParse(distanta.Replace("max", ""), out offset) != false)
                            {
                                offset = sect[sect.Count - 1].Offset + offset;
                                Punct3D p = new Punct3D();
                                p.KM = sect[0].KM;
                                p.Offset = offset;
                                p.Z = -999;
                                p.D = descr;
                                //ed.WriteMessage(" ---->");
                                listanoua.Add(p);
                                //ed.WriteMessage(" " + offset.ToString());
                            }
                            else
                            {
                                ed.WriteMessage("Offset {0} is invalid and will be ignored.", distanta);
                                continue;
                            }
                        }
                        else
                        {
                            if (double.TryParse(distanta, out offset) != false)
                            {
                                Punct3D p = new Punct3D();
                                p.KM = sect[0].KM;
                                p.Offset = offset;
                                p.Z = -999;
                                p.D = descr;
                                //ed.WriteMessage(" ---->");
                                listanoua.Add(p);
                                //ed.WriteMessage(" " + listanoua[listanoua.Count -1].Offset);
                            }
                            else
                            {
                                ed.WriteMessage("Offset {0} is invalid and will be ignored.", distanta);
                                continue;
                            }
                        }
                    }
                }
            }
            ed.WriteMessage("\nTotally {0} new points will be calculated, ", listanoua.Count);
            
            ///DUPA OBTINEREA LISTEI DE PUNCTE NOI
            // Sortarea listei
            listanoua.Sort(
                delegate(Punct3D p1, Punct3D p2)
                {
                    if (p1.KM.CompareTo(p2.KM) == 0) return p1.Offset.CompareTo(p2.Offset);
                    else return p1.KM.CompareTo(p2.KM);
                }
            );
            // Copierea ei pe sectiuni si apoi golirea
            List<String3D> sectiuninoi = new List<String3D>();
            String3D sectiunenoua = new String3D();
            sectiunenoua.Add(listanoua[0]);
            for (int i = 1; i < listanoua.Count; i++)
            {
                if (listanoua[i].KM != listanoua[i - 1].KM)
                {
                    sectiuninoi.Add(sectiunenoua);
                    sectiunenoua = new String3D();
                }
                sectiunenoua.Add(listanoua[i]);
            }
            sectiuninoi.Add(sectiunenoua);
            listanoua = new String3D();
            ed.WriteMessage("arranged in {0} sections between km {1} and km {2}",
                sectiuninoi.Count, sectiuninoi[0][0].KM, sectiuninoi[sectiuninoi.Count - 1][0].KM);


            //Parcurgerea sectiunilor noi
            foreach (String3D SN in sectiuninoi)
            {
                String3D SV = new String3D();
                try
                {
                    SV = sectiuni.Find(s => s[0].KM == SN[0].KM);
                    if (SV.Count < 2 || SV[SV.Count - 1].Offset - SV[0].Offset == 0) throw new SystemException("Insufficien points");
                    //ed.WriteMessage("\nInterpolating on section {0}", SV[0].KM);
                }
                catch
                {
                    ed.WriteMessage("\nInssuficient points at km {0} in the source station-offset file! Skipping section.");
                    continue;
                }

                foreach (Punct3D PN in SN)
                {
                    //punct gasit in sectiunea existenta
                    try { PN.Z = SV.Find(punct => punct.Offset == PN.Offset).Z; }
                    catch
                    {
                        //extrapolare stanga
                        if (PN.Offset < SV[0].Offset)
                            if (PKR2.StringResult == "Yes")
                                PN.Z = SV[1].Z - (SV[1].Z - SV[0].Z) * (SV[1].Offset - PN.Offset) / (SV[1].Offset - SV[0].Offset);
                            else continue;
                        //extrapolare dreapta
                        else if (PN.Offset > SV[SV.Count - 1].Offset)
                            if (PKR2.StringResult == "Yes")
                                PN.Z = SV[SV.Count - 2].Z + (SV[SV.Count - 1].Z - SV[SV.Count - 2].Z) *
                                    (PN.Offset - SV[SV.Count - 2].Offset) / (SV[SV.Count - 1].Offset - SV[SV.Count - 2].Offset);
                            else continue;
                        //interpolare obisnuita
                        else
                        {
                            Punct3D P2 = SV.Find(punct => punct.Offset > PN.Offset);
                            Punct3D P1 = SV[SV.IndexOf(P2) - 1];
                            PN.Z = P1.Z + (P2.Z - P1.Z) * (PN.Offset - P1.Offset) / (P2.Offset - P1.Offset);
                        }
                    }
                    listanoua.Add(PN);
                }
            }


            ///DUPA OBTINEREA PUNCTELOR NOI CU COTE INTERPOLATE
            // Daca s-a optat pentru includerea tuturor punctelor, se adauga cele vechi la lista noua si se sorteaza;
            
            if (PKR.StringResult == "All")
            {
                foreach (Punct3D p in listapuncte) listanoua.Add(p);
            }
            listanoua.Sort(
                delegate(Punct3D p1, Punct3D p2)
                {
                    if (p1.KM.CompareTo(p2.KM) == 0) return p1.Offset.CompareTo(p2.Offset);
                    else return p1.KM.CompareTo(p2.KM);
                }
            );
            
            string endfile = startfile.Insert(startfile.LastIndexOf('.'), "-LINTERP");
            StreamWriter scriitor = new StreamWriter(endfile, true);
            foreach (Punct3D p in listanoua)
            {
                scriitor.WriteLine(p.toString(Punct3D.Format.KmOZD, Punct3D.DelimitedBy.Comma, 4, false));
            }
            scriitor.Dispose();
            ed.WriteMessage("\nThe result file has {0} points.", listanoua.Count);
        }


        //Comanda pentru extinderea/taierea unui set de sectiuni uzand de un al doilea asemenea set
        [CommandMethod("TRIMEXT")]
        public void trimext()
        {
            Document acadDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acadDoc.Editor;
            //Se creaza lista de puncte noi, ale caro cote vor fi calculate prin interpolare
            String3D listanoua = new String3D();

            #region SELECTIA FISIERELOR
            //selectia fisierului STO cu date de modificat
            ed.WriteMessage("\nData source file selection:");
            PromptOpenFileOptions PFO = new PromptOpenFileOptions("\nSelect station-offset file: ");
            PFO.Filter = "Text file (*.txt)|*.txt|Comma separated file (*.csv)|*.csv";
            PromptFileNameResult PFR = ed.GetFileNameForOpen(PFO);
            if (PFR.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nInvalid file selection! Aborting.");
                return;
            }
            string startfile = PFR.StringResult;

            //selectia fisierului STO tinta
            ed.WriteMessage("\nTarget data file selection:");
            PFR = ed.GetFileNameForOpen(PFO);
            if (PFR.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nInvalid file selection! Aborting.");
                return;
            }
            string targetfile = PFR.StringResult;
            #endregion

            #region PROCESAREA FISIERELOR
            ed.WriteMessage("\nProcessing source data file: ");
            List<String3D> StartSections = ProcessSTOfile(startfile);
            ed.WriteMessage("\nProcessing target data file: ");
            List<String3D> TargetSections = ProcessSTOfile(targetfile);
            ed.WriteMessage("\nProcessing output file: ");
            List<String3D> FinalSections = new List<String3D>();

            foreach (String3D XS in StartSections)
            {
                if (TargetSections.Exists(X => X[0].KM == XS[0].KM))
                {
                    FinalSections.Add(StringUtil.StringUtil.XStrimextend(XS, TargetSections.Find(X => X[0].KM == XS[0].KM)));
                }
                else
                {
                    FinalSections.Add(XS);
                }
            }
            #endregion

            #region SCRIEREA FISIERULUI REZULTAT
            string endfile = startfile.Insert(startfile.LastIndexOf('.'), "-TRIMEXT");
            StreamWriter scriitor = new StreamWriter(endfile, true);
            int nrPuncte = 0;
            foreach (String3D XS in FinalSections)
            {
                foreach (Punct3D P in XS)
                {
                    scriitor.WriteLine(P.toString(Punct3D.Format.KmOZD, Punct3D.DelimitedBy.Comma, 4, false));
                }
                nrPuncte += XS.Count;
            }
            scriitor.Dispose();
            ed.WriteMessage("\n{0} points were written to the output file, arranged in {1} sections from km {2} to km {3}",
                nrPuncte, FinalSections.Count, FinalSections[0][0].KM, FinalSections[FinalSections.Count - 1][0].KM);
            #endregion
        }
        
        private static List<String3D> ProcessSTOfile(string file)
        {
            Document acadDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acadDoc.Editor;
            //Se creaza lista de puncte noi, ale caro cote vor fi calculate prin interpolare
            String3D listanoua = new String3D();

            //Citirea tuturor punctelor din fisierul sursa si sortarea lor dupa km si offset
            String3D listastart = new String3D();
            String3D listatinta = new String3D();
            listastart.ImportPoints(file, Punct3D.Format.KmOZD, Punct3D.DelimitedBy.Comma);
            if (listastart.Count < 2)
            {
                ed.WriteMessage("\nNot enough points were read from the selected file! Aborting.");
                return null;
            }
            listastart.Sort(
                delegate(Punct3D p1, Punct3D p2)
                {
                    if (p1.KM.CompareTo(p2.KM) == 0) return p1.Offset.CompareTo(p2.Offset);
                    else return p1.KM.CompareTo(p2.KM);
                }
            );


            //Impartirea punctelor pe sectiuni
            List<String3D> sectiuni = new List<String3D>();
            String3D sectiune = new String3D();
            sectiune.Add(listastart[0]);
            for (int i = 1; i < listastart.Count; i++)
            {
                if (listastart[i].KM != listastart[i - 1].KM)
                {
                    sectiuni.Add(sectiune);
                    sectiune = new String3D();
                }
                sectiune.Add(listastart[i]);
            }
            sectiuni.Add(sectiune);
            ed.WriteMessage("\n{0} points were read from the source file, arranged in {1} sections from km {2} to km {3}",
                listastart.Count, sectiuni.Count, sectiuni[0][0].KM, sectiuni[sectiuni.Count - 1][0].KM);
            return sectiuni;
        }
    }

    public class DwgUtil : IExtensionApplication
        {
            #region IExtensionApplication Members
            public void Initialize()
            {
                //throw new NotImplementedException();
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\nLoaded AcadUTIL - small routines created by Ovidiu Candrea <ovidiucandrea@yahoo.com>\nFor help run command HELPUTIL");
            }

            public void Terminate()
            {
                //throw new NotImplementedException();
            }
            #endregion

            [CommandMethod("DES")]
            public void deschide()
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                string caleDesen = "";

                //Cauta fisierul de configurare in calea implicita sau cere calea
                string cale;
                cale = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Optiuni.cfg";
                cale = Path.GetFullPath(cale);
                if (File.Exists(cale) == false)
                {
                    //PromptOpenFileOptions prFileOpt = new PromptOpenFileOptions("Select Configuration File");
                    //prFileOpt.Filter = "Configuration File (*.cfg)";
                    //prFileOpt.InitialFileName = "Optiuni.cfg";
                    try
                    {
                        using (StreamWriter scriitor = new StreamWriter(cale))
                        {
                            //cale = ed.GetFileNameForOpen(prFileOpt).StringResult;
                            //cale = Path.GetFullPath(cale);
                            scriitor.Write("\r\n");
                        }
                    }
                    catch
                    {
                        ed.WriteMessage("Unsuccessful selection of Configuration File");
                    }
                }
                Ovidiu.x64.General.Configurator config = new Ovidiu.x64.General.Configurator(cale);

                //Optiunile afisate la cererea numelui desenului de deschis
                PromptKeywordOptions prOpt = new PromptKeywordOptions("Open Drawing:");
                prOpt.Keywords.Add("List");
                prOpt.Keywords.Add("New");
                prOpt.AppendKeywordsToMessage = true;
                prOpt.AllowArbitraryInput = true;


                //Cererea rezultatului in linia de comanda
                PromptResult prRes = ed.GetKeywords(prOpt);

                //Analiza rezultatului
                if (prRes.Status == PromptStatus.OK)
                {
                    ed.WriteMessage(prRes.StringResult);
                    string[] optiuni = config.citireConf("@deseneACAD");
                    switch (prRes.StringResult)
                    {
                        case "List":

                            ed.WriteMessage("\r\n");
                            foreach (string optiune in optiuni)
                            {
                                //ed.WriteMessage(optiune.Substring(optiune.IndexOf("=") + 1) + "\r\n");
                                ed.WriteMessage(optiune + "\r\n");
                            }
                            break;
                        case "New":
                            try
                            {
                                PromptFileNameResult PFO = ed.GetFileNameForOpen("\nOpen Drawing:");
                                if (PFO.Status != PromptStatus.OK) return;
                                caleDesen = PFO.StringResult;
                                PromptStringOptions prStrOpt = new PromptStringOptions("\nSpecify the key for the new drawing: ");
                                //prStrOpt.DefaultValue = "NewKey";
                                string cheie = ed.GetString(prStrOpt).StringResult;
                                //Evitarea conflictului de nume
                                while (cheie == "List" || cheie == "New")
                                {
                                    cheie = ed.GetString("\nList and New are not valid keys, please select another key: ").StringResult;
                                }
                                string[] optNoua = new string[1];
                                optNoua[0] = cheie + "=" + caleDesen;
                                if (caleDesen != "") caleDesen = Path.GetFullPath(caleDesen);
                                config.scriereConf("@deseneACAD", optNoua);
                                Application.DocumentManager.MdiActiveDocument = Application.DocumentManager.Open(caleDesen, false);
                                caleDesen = "";
                            }
                            catch
                            {
                            }
                            break;
                        default:
                            try
                            {
                                bool cheieGasita = false;
                                foreach (string optiune in optiuni)
                                {
                                    if (optiune.StartsWith(prRes.StringResult + "="))
                                    {
                                        cheieGasita = true;
                                        caleDesen = Path.GetFullPath(optiune.Replace(prRes.StringResult + "=", ""));
                                    }
                                }
                                if (cheieGasita)
                                {
                                    Application.DocumentManager.MdiActiveDocument = Application.DocumentManager.Open(caleDesen, false);
                                    caleDesen = "";
                                }
                                else
                                {
                                    ed.WriteMessage("\nKey not found; try a new one!");
                                }
                            }
                            catch
                            {
                            }
                            break;
                    }
                }
            }

            [CommandMethod("ATAS")]
            public void ataseaza()
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                string caleDesen = "";

                //Cauta fisierul de configurare in calea implicita sau cere calea
                string cale;
                cale = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Optiuni.cfg";
                cale = Path.GetFullPath(cale);
                if (File.Exists(cale) == false)
                {
                    //PromptOpenFileOptions prFileOpt = new PromptOpenFileOptions("Select Configuration File");
                    //prFileOpt.Filter = "Configuration File (*.cfg)";
                    //prFileOpt.InitialFileName = "Optiuni.cfg";
                    try
                    {
                        using (StreamWriter scriitor = new StreamWriter(cale))
                        {
                            //cale = ed.GetFileNameForOpen(prFileOpt).StringResult;
                            //cale = Path.GetFullPath(cale);
                            scriitor.Write("\r\n");
                        }
                    }
                    catch
                    {
                        ed.WriteMessage("Unsuccessful selection of Configuration File");
                    }
                }
                Ovidiu.x64.General.Configurator config = new Ovidiu.x64.General.Configurator(cale);

                //Optiunile afisate la cererea numelui desenului de deschis
                PromptKeywordOptions prOpt = new PromptKeywordOptions("Attach Drawing:");
                prOpt.Keywords.Add("List");
                prOpt.Keywords.Add("New");
                prOpt.AppendKeywordsToMessage = true;
                prOpt.AllowArbitraryInput = true;


                //Cererea rezultatului in linia de comanda
                PromptResult prRes = ed.GetKeywords(prOpt);

                //Analiza rezultatului
                if (prRes.Status == PromptStatus.OK)
                {
                    ed.WriteMessage(prRes.StringResult);
                    string[] optiuni = config.citireConf("@deseneACAD");
                    switch (prRes.StringResult)
                    {
                        case "List":

                            ed.WriteMessage("\r\n");
                            foreach (string optiune in optiuni)
                            {
                                //ed.WriteMessage(optiune.Substring(optiune.IndexOf("=") + 1) + "\r\n");
                                ed.WriteMessage(optiune + "\r\n");
                            }
                            break;
                        case "New":
                            try
                            {
                                PromptFileNameResult PFO = ed.GetFileNameForOpen("\nOpen Drawing:");
                                if (PFO.Status != PromptStatus.OK) return;
                                caleDesen = PFO.StringResult;

                                PromptStringOptions prStrOpt = new PromptStringOptions("\nSpecify the key for the new drawing: ");
                                //prStrOpt.DefaultValue = "NewKey";

                                Point3d pctInsertie = new Point3d(0, 0, 0);
                                PromptPointOptions prPtOpt = new PromptPointOptions("\nSpecify a point for insertion: ");
                                prPtOpt.AllowNone = true;
                                PromptPointResult prPtRes = ed.GetPoint(prPtOpt);
                                if (prPtRes.Status == PromptStatus.OK) pctInsertie = prPtRes.Value;

                                string cheie = ed.GetString(prStrOpt).StringResult;
                                //Evitarea conflictului de nume
                                while (cheie == "List" || cheie == "New")
                                {
                                    cheie = ed.GetString("\nList and New are not valid keys, please select another key: ").StringResult;
                                }
                                string[] optNoua = new string[1];
                                optNoua[0] = cheie + "=" + caleDesen;
                                if (caleDesen != "") caleDesen = Path.GetFullPath(caleDesen);
                                config.scriereConf("@deseneACAD", optNoua);
                                atasXRef(caleDesen, pctInsertie); //Atasarea referintei externe
                                caleDesen = "";
                            }
                            catch
                            {
                            }
                            break;
                        default:
                            try
                            {
                                bool cheieGasita = false;
                                foreach (string optiune in optiuni)
                                {
                                    if (optiune.StartsWith(prRes.StringResult + "="))
                                    {
                                        cheieGasita = true;
                                        caleDesen = Path.GetFullPath(optiune.Replace(prRes.StringResult + "=", ""));
                                    }
                                }
                                if (cheieGasita)
                                {
                                    Point3d pctInsertie = new Point3d(0, 0, 0);
                                    PromptPointOptions prPtOpt = new PromptPointOptions("\nSpecify a point for insertion: ");
                                    prPtOpt.AllowNone = true;
                                    PromptPointResult prPtRes = ed.GetPoint(prPtOpt);
                                    if (prPtRes.Status == PromptStatus.OK) pctInsertie = prPtRes.Value;

                                    atasXRef(caleDesen, pctInsertie); //Atasarea referintei externe
                                    caleDesen = "";
                                }
                                else
                                {
                                    ed.WriteMessage("\nKey not found; try a new one!");
                                }
                            }
                            catch
                            {
                            }
                            break;
                    }
                }
            }

            [CommandMethod("+Cod")]
            public void plusCod()
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                string[] optiuni = new string[1];
                optiuni[0] = Path.GetFullPath(Application.DocumentManager.MdiActiveDocument.Name);

                //Cauta fisierul de configurare in calea implicita sau cere calea
                string cale;
                cale = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Optiuni.cfg";
                cale = Path.GetFullPath(cale);
                if (File.Exists(cale) == false)
                {
                    //PromptOpenFileOptions prFileOpt = new PromptOpenFileOptions("Select Configuration File");
                    //prFileOpt.Filter = "Configuration File (*.cfg)";
                    //prFileOpt.InitialFileName = "Optiuni.cfg";
                    try
                    {
                        using (StreamWriter scriitor = new StreamWriter(cale))
                        {
                            //cale = ed.GetFileNameForOpen(prFileOpt).StringResult;
                            //cale = Path.GetFullPath(cale);
                            scriitor.Write("\r\n");
                        }
                    }
                    catch
                    {
                        ed.WriteMessage("\nUnsuccessful selection of Configuration File");
                    }
                }
                Ovidiu.x64.General.Configurator config = new Ovidiu.x64.General.Configurator(cale);

                //Atasarea codului pentru desenul curent la fisierul de configurare
                try
                {
                    string cod = ed.GetString("\nEnter the key for further usage of the active drawing: ").StringResult;
                    config.scriereConf(cod, optiuni);
                }
                catch
                {
                    ed.WriteMessage("\nUnsuccessful writing of Configuration File");
                }
            }

            //Transparent Commands
            [CommandMethod("'plv",CommandFlags.Transparent)]
            public void addpolysector()
            {
                try
                {
                    Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                    Database db = HostApplicationServices.WorkingDatabase;
                    Editor ed = acadDoc.Editor;

                    //Check for command compatibility (PLINE and 3DPOLY addpoints, no active command starts a new PLINE)
                    string runningCMD = Application.GetSystemVariable("CMDNAMES") as string;
                    if (!runningCMD.StartsWith("PLINE") && !runningCMD.StartsWith("3DPOLY") && !runningCMD.StartsWith("PLV"))
                    {
                        ed.WriteMessage("\n'plv only runs from within PLINE or 3DPOLY commands! {0}");
                        return;
                    }


                    //Select polyline from which to add vertices
                    PromptEntityOptions PrEntOpt = new PromptEntityOptions("\nSelect lightweight polyline from which to add vertices: ");
                    PrEntOpt.SetRejectMessage("\nObject is not a lightweight polyline!");
                    PrEntOpt.AddAllowedClass(typeof(Polyline), true);
                    PrEntOpt.AllowObjectOnLockedLayer = true;
                    PromptEntityResult PrEntRes = ed.GetEntity(PrEntOpt);
                    if (PrEntRes.Status != PromptStatus.OK)
                    {
                        return;
                    }

                    //Select vertices to add to current polyline
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        Polyline srcPoly = (Polyline)trans.GetObject(PrEntRes.ObjectId, OpenMode.ForRead);
                        ed.WriteMessage("\nSelected polyline is of type {0} and has {1} vertices", srcPoly.GetType().Name, srcPoly.NumberOfVertices);
                        //PlanarEntity plan = srcPoly.GetPlane();

                        ////Not required in the original version
                        //BlockTable BT = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                        //BlockTableRecord BTR = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        ////

                        Point3d StartPnt = new Point3d();
                        Point3d EndPnt = new Point3d();

                        //Select starting point
                        PromptPointOptions PrPntOpt = new PromptPointOptions("\nSelect start vertex: ");
                        PrPntOpt.Keywords.Add("Start");
                        PrPntOpt.Keywords.Add("End");
                        PrPntOpt.Keywords.Add("Number");
                        PrPntOpt.AppendKeywordsToMessage = true;

                        PromptPointResult PrPntRes = ed.GetPoint(PrPntOpt);
                        switch (PrPntRes.Status)
                        {
                            case PromptStatus.Keyword:
                                switch (PrPntRes.StringResult)
                                {
                                    case "Start":
                                        StartPnt = srcPoly.GetPoint3dAt(0);
                                        break;
                                    case "End":
                                        StartPnt = srcPoly.GetPoint3dAt(srcPoly.NumberOfVertices - 1);
                                        break;
                                    case "Number":
                                        PromptIntegerOptions PrIntOpt = new PromptIntegerOptions("\nSpecify the start vertex number: ");
                                        PrIntOpt.AllowArbitraryInput = false;
                                        PrIntOpt.LowerLimit = 1;
                                        PrIntOpt.UpperLimit = srcPoly.NumberOfVertices;
                                        StartPnt = srcPoly.GetPoint3dAt(ed.GetInteger(PrIntOpt).Value - 1);
                                        break;
                                }
                                break;
                            case PromptStatus.OK:
                                StartPnt = srcPoly.GetClosestPointTo(PrPntRes.Value, true);
                                break;
                            default:
                                return;
                        }

                        //Select end point
                        PrPntOpt.Message = "\nSelect end vertex: ";
                        PrPntRes = ed.GetPoint(PrPntOpt);
                        switch (PrPntRes.Status)
                        {
                            case PromptStatus.Keyword:
                                switch (PrPntRes.StringResult)
                                {
                                    case "Start":
                                        EndPnt = srcPoly.GetPoint3dAt(0);
                                        break;
                                    case "End":
                                        EndPnt = srcPoly.GetPoint3dAt(srcPoly.NumberOfVertices - 1);
                                        break;
                                    case "Number":
                                        PromptIntegerOptions PrIntOpt = new PromptIntegerOptions("\nSpecify the end vertex number: ");
                                        PrIntOpt.AllowArbitraryInput = false;
                                        PrIntOpt.LowerLimit = 1;
                                        PrIntOpt.UpperLimit = srcPoly.NumberOfVertices;
                                        StartPnt = srcPoly.GetPoint3dAt(ed.GetInteger(PrIntOpt).Value - 1);
                                        break;
                                }
                                break;
                            case PromptStatus.OK:
                                EndPnt = srcPoly.GetClosestPointTo(PrPntRes.Value, true);
                                break;
                            default:
                                return;
                        }

                        //Obtain the string with point coordinates to be added
                        int StartNr = -1;
                        int EndNr = -1;
                        List<string> substring = new List<string>();
                        //List<Point2d> substring = new List<Point2d>();
                        for (int i = 0; i < srcPoly.NumberOfVertices; i++)
                        {
                            if (srcPoly.GetPoint3dAt(i).IsEqualTo(StartPnt))
                            {
                                StartNr = i;
                            }
                            if (srcPoly.GetPoint3dAt(i).IsEqualTo(EndPnt))
                            {
                                EndNr = i;
                            }
                            if (StartNr + EndNr != -2 && i * Math.Sign(Math.Min(EndNr,StartNr)) <= Math.Max(StartNr, EndNr))
                            {
                                substring.Add(string.Format("{0},{1}", srcPoly.GetPoint3dAt(i).X, srcPoly.GetPoint3dAt(i).Y));
                                //substring.Add(srcPoly.GetPoint2dAt(i));
                            }
                        }
                        if (StartNr > EndNr)
                        {
                            substring.Reverse();
                        }

                        ////Obtain the current plane
                        //try
                        //{
                        //    UcsTable UCST = (UcsTable)trans.GetObject(db.UcsTableId, OpenMode.ForRead);
                        //    UcsTableRecord UCSR = (UcsTableRecord)trans.GetObject(UCST[(string)Application.GetSystemVariable("UCSNAME")], OpenMode.ForRead);
                        //    UCSR.vect
                        //}
                        //catch
                        //{
                        //}

                        ////Generate the result polyline
                        //Polyline rezpoly = new Polyline(substring.Count);
                        //int plinewid = 0;
                        //try 
                        //{
                        //    plinewid = (int)Application.GetSystemVariable("PLINEWID");
                        //}
                        //catch
                        //{
                        //}
                        
                        //for (int i = 0; i < substring.Count; i++)
                        //{
                        //    rezpoly.AddVertexAt(i, substring[i], 0, plinewid, plinewid);
                        //}

                        //BTR.AppendEntity(rezpoly);
                        //trans.AddNewlyCreatedDBObject(rezpoly, true);
                        //trans.Commit();



                        //Obtain the string to be passed to the 'PLINE' command
                        string comanda = string.Empty;
                        foreach (string punct in substring)
                        {
                            comanda += punct + "\n";
                        }
                        //Check if 'PLINE' command is running
                        if (runningCMD.StartsWith("PLV"))
                        {
                            comanda = "PLINE " + comanda;
                        }

                        //ed.WriteMessage(comanda);
                        acadDoc.SendStringToExecute(comanda, true, false, true);
                    }
                }
                catch
                {
                }
            }

            //New Palette
            //[CommandMethod("TestPalette")]
            //public void DoIt()
            //{
            //    PaletteSet ps = new PaletteSet("BrowserPalette");
                
            //}

            public void atasXRef(string cale, Point3d pctInsertie)
            {
                Database db = HostApplicationServices.WorkingDatabase;
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {

                    try
                    {
                        //BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        ObjectId xRefId = db.AttachXref(cale, Path.GetFileNameWithoutExtension(cale));
                        BlockReference br = new BlockReference(pctInsertie, xRefId);
                        btr.AppendEntity(br);
                        trans.AddNewlyCreatedDBObject(br, true);
                        trans.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(ex.ToString());
                    }
                }
            }

        }

    public class SectUtil
        {
            //Comanda nefinalizata!
            [CommandMethod("EXSEC")]
            public void exsec()
            {
                Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                Database db = acadDoc.Database;
                Editor ed = acadDoc.Editor;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);


                    Autodesk.AutoCAD.Windows.OpenFileDialog fileDia = new Autodesk.AutoCAD.Windows.OpenFileDialog
                        ("Select file", "Point Export", "csv", "Select file",
                        Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension);
                    string cale = null;
                    try
                    {
                        cale = fileDia.Filename;
                    }
                    catch (System.Exception)
                    {


                    }

                    if (cale == null) return;


                    //System.Windows.Forms.DialogResult diaRes = fileDia.ShowDialog();
                    //if (diaRes != System.Windows.Forms.DialogResult.OK) return;
                    //string cale = diaRes.ToString();

                    //StreamWriter scriitor = new StreamWriter(cale);

                    try
                    {
                        while (true)
                        {
                            //Selectia punctului de referinta
                            Point3d pRef = Point3d.Origin;
                            PromptPointResult PrPointRes = ed.GetPoint("\nSelect reference point");
                            if (PrPointRes.Status == PromptStatus.OK)
                            {
                                pRef = PrPointRes.Value;
                                ed.WriteMessage("\nReference point selected");
                            }

                            ////Selectia cotei de referinta
                            //double cotaRef = -999;
                            //PromptEntityOptions PrEntOpt = new PromptEntityOptions("\nSelect Text containing reference level");
                            //PrEntOpt.AddAllowedClass(typeof(MText), true);
                            //PrEntOpt.AddAllowedClass(typeof(DBText), true);
                            //PrEntOpt.AllowNone = false;
                            //PromptEntityResult PrEntRes = ed.GetEntity(PrEntOpt);
                            //if (PrEntRes.Status == PromptStatus.OK)
                            //{
                            //    ObjectId ObjId = PrEntRes.ObjectId;
                            //    DBObject Obj = trans.GetObject(ObjId, OpenMode.ForRead);

                            //    if (Obj as MText != null)
                            //    {
                            //        MText mText = (MText)Obj;
                            //        cotaRef = double.Parse(mText.Text);
                            //    }
                            //    else
                            //    {
                            //        DBText Text = (DBText)Obj;
                            //        cotaRef = double.Parse(Text.TextString);
                            //    }
                            //}

                            ////Selectia kilometrajului
                            //double km = -999;
                            //PrEntOpt = new PromptEntityOptions("\nSelect Text containing reference level");
                            //PrEntOpt.AddAllowedClass(typeof(MText), true);
                            //PrEntOpt.AddAllowedClass(typeof(DBText), true);
                            //PrEntOpt.AllowNone = false;
                            //PrEntRes = ed.GetEntity(PrEntOpt);
                            //if (PrEntRes.Status == PromptStatus.OK)
                            //{
                            //    ObjectId ObjId = PrEntRes.ObjectId;
                            //    DBObject Obj = trans.GetObject(ObjId, OpenMode.ForRead);

                            //    if (Obj as MText != null)
                            //    {
                            //        MText mText = (MText)Obj;
                            //        km = double.Parse(mText.Text);
                            //    }
                            //    else
                            //    {
                            //        DBText Text = (DBText)Obj;
                            //        km = double.Parse(Text.TextString);
                            //    }
                            //}

                            double cotaRef = -999;
                            double km = -999;
                            cotaRef = ed.GetDouble("\nEnter reference point level").Value;
                            km = ed.GetDouble("\nEnter chainage").Value;

                            //Selectia poliliniei ce trebuie citita
                            Polyline poly = null;
                            PromptEntityOptions PrEntOpt = new PromptEntityOptions("\nSelect Text containing reference level");
                            PrEntOpt.AddAllowedClass(typeof(Polyline), true);
                            PrEntOpt.AllowNone = false;
                            PromptEntityResult PrEntRes = ed.GetEntity(PrEntOpt);
                            if (PrEntRes.Status == PromptStatus.OK)
                            {
                                ObjectId ObjId = PrEntRes.ObjectId;
                                poly = (Polyline)trans.GetObject(ObjId, OpenMode.ForRead);
                            }

                            //Citirea poliliniei si convertirea punctelor
                            if (pRef != Point3d.Origin && cotaRef != -999 && km != -999 && poly != null)
                            {
                                String3D listaPuncte = new String3D();
                                List<Point3d> pointList = new List<Point3d>();
                                for (int i = 0; i < poly.NumberOfVertices; i++)
                                {
                                    pointList.Add(poly.GetPoint3dAt(i));
                                }

                                List<Point3d> convertedPoints = Transfom2Param(pointList, pRef.X, pRef.Y, 0, cotaRef);
                                ed.WriteMessage("\n{0} points read and converted!", convertedPoints.Count);

                                foreach (Point3d p in convertedPoints)
                                {
                                    Punct3D punct = new Punct3D();
                                    punct.KM = km;
                                    punct.Offset = p.X;
                                    punct.Z = p.Y;
                                    listaPuncte.Add(punct);
                                }

                                listaPuncte.ExportPoints(cale, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 4, false);
                            }


                        }
                    }
                    catch (System.Exception)
                    {

                    }
                }
            }

            //Comanda pentru citirea sistemelor de referinta ale sectiunilor si scrierea lor in fisier *.SCS
            [CommandMethod("SCS")]
            public void scs()
            {
                Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = acadDoc.Editor;
                Database db = HostApplicationServices.WorkingDatabase;

                //Se trece la ucs global
                ed.CurrentUserCoordinateSystem = new Matrix3d(new double[16]{
    1.0, 0.0, 0.0, 0.0,
    0.0, 1.0, 0.0, 0.0,
    0.0, 0.0, 1.0, 0.0,
    0.0, 0.0, 0.0, 1.0});

                //Gasirea caii documentului curent
                string caleDocAcad = acadDoc.Database.Filename;
                //Daca fisierul este unul nou, ne salvat inca, calea apartine sablonului folosit pentru acesta.
                //Se verifica daca fisierul este unul sablon, se atentioneaza utilizatorul si se paraseste programul.
                if (caleDocAcad.EndsWith(".dwt") == true)
                {
                    ed.WriteMessage("\nThe current drawing is a template file (*.dwt). Exiting program! ");
                    return;
                }
                caleDocAcad = HostApplicationServices.Current.FindFile(acadDoc.Name, acadDoc.Database, FindFileHint.Default);
                //string caleDocAcad = HostApplicationServices.Current.FindFile(acadDoc.Name, acadDoc.Database, FindFileHint.Default);

                //Citire registru fisiere cu sisteme de coordonate ale sectiunilor ANULAT
                //Ovidiu.x64.General.Configurator config = GetConfig();
                //string[] registru = config.citireConf("@SectionCSfiles");

                //string caleFisCS = string.Empty;
                FileInfo fisier = null;
                #region Verificare fisier existent si solicitare procedura inlocuire/adnotare/anulare
                //foreach (string cale in registru)
                //{
                //    if (caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) == cale.Remove(cale.LastIndexOf('.')))
                //    {
                //        caleFisCS = cale;
                //        continue;
                //    }
                //}
                fisier = new FileInfo(caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + ".SCS");
                //if (caleFisCS != string.Empty && new FileInfo(caleDocAcad).Exists == true)
                if (fisier.Exists)
                {
                    PromptKeywordOptions PrKeyOpt = new PromptKeywordOptions("");
                    PrKeyOpt.Message = "The Section Coordinate System File already exists! ";
                    PrKeyOpt.Keywords.Add("Overwrite");
                    PrKeyOpt.Keywords.Add("Append");
                    PrKeyOpt.Keywords.Add("eXit");
                    PrKeyOpt.AppendKeywordsToMessage = true;
                    PrKeyOpt.AllowArbitraryInput = false;
                    PromptResult PrKeyRes = ed.GetKeywords(PrKeyOpt);
                    if (PrKeyRes.Status == PromptStatus.OK)
                    {
                        switch (PrKeyRes.StringResult)
                        {
                            case "Overwrite":
                                //Se salveaza o copie redenumita a fisierului cu sisteme de coordonate si se furnizeaza calea
                                //fisier = new FileInfo(caleFisCS);
                                FileInfo backup = fisier.CopyTo(fisier.FullName.Remove(fisier.FullName.LastIndexOf('.')) + "~BACKUP.BCS", true);
                                fisier.Delete();
                                StreamWriter scriitor = new StreamWriter(fisier.FullName);
                                ed.WriteMessage("\nThe file was overwritten! A backup of the previous file was saved at: \n{0}",
                                    backup.FullName);
                                scriitor.Close();
                                scriitor.Dispose();
                                break;

                            case "Append":
                                //fisier = new FileInfo(caleFisCS);
                                break;

                            case "Exit":
                            default:
                                ed.WriteMessage("\nExiting Program! ");
                                return;
                        }
                    }
                }
                #endregion

                //Daca registrul nu contine niciun fisier cu sisteme de coordonate ale sectiunilor sau fisierul din registru nu exista
                else
                {
                    //fisier = new FileInfo(caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + ".SCS");
                    StreamWriter scriitor = new StreamWriter(fisier.FullName);
                    scriitor.Close();
                    scriitor.Dispose();
                }

                String3D listaCS = new String3D();
                List<double> listaKm = new List<double>();
                #region Citirea fisierului cu sisteme de coordonate ale sectiunilor si afisarea pozitiilor gasite
                listaCS.ImportPoints(fisier.FullName, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma);
                if (listaCS.Count != 0)
                {
                    ed.WriteMessage("\nThe following chainages have already been found in the SCS file: ");
                    foreach (Punct3D CS in listaCS)
                    {
                        ed.WriteMessage("\n{0}", CS.toString(Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 4, true));
                        listaKm.Add(CS.KM);
                    }
                }
                #endregion



                #region Inregistrarea coordonatelor de referinta ale sectiunilor
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                    PromptEntityOptions PrEntKmOpt = new PromptEntityOptions("");
                    PrEntKmOpt.Message = "\nIndicate Chainage (Text or Mtext)";
                    PrEntKmOpt.Keywords.Add("Value");
                    PrEntKmOpt.Keywords.Add("eXit");
                    PrEntKmOpt.AppendKeywordsToMessage = true;
                    PrEntKmOpt.SetRejectMessage("\nThe selected object is not of required type! ");
                    PrEntKmOpt.AddAllowedClass(typeof(DBText), true);
                    PrEntKmOpt.AddAllowedClass(typeof(MText), true);
                    PrEntKmOpt.AllowNone = true;

                    bool exit = false;
                    while (exit == false)
                    {
                        string textKm = string.Empty;

                        //if (listaKm.Count != 0)
                        //{
                        //    PrEntKmOpt.Keywords.Default = (listaKm[listaKm.Count - 1] + 25).ToString();
                        //}
                        PromptEntityResult PrEntKmRes = ed.GetEntity(PrEntKmOpt);
                        if (PrEntKmRes.Status == PromptStatus.OK)
                        {
                            ObjectId objID = PrEntKmRes.ObjectId;
                            Autodesk.AutoCAD.DatabaseServices.DBObject obj = trans.GetObject(objID, OpenMode.ForRead);
                            if (obj is DBText)
                            {
                                DBText text = (DBText)obj;
                                textKm = text.TextString;
                            }
                            if (obj is MText)
                            {
                                MText mtext = (MText)obj;
                                textKm = mtext.Contents;
                            }
                        }
                        else if (PrEntKmRes.Status == PromptStatus.Cancel)
                        {
                            ed.WriteMessage("\nCommand aborted. Exiting!");
                            goto final;
                        }
                        else if (PrEntKmRes.StringResult != string.Empty)
                        {
                            if (PrEntKmRes.StringResult == "eXit")
                            {
                                ed.WriteMessage("\nCommand aborted. Exiting!");
                                goto final;
                            }
                            else if (PrEntKmRes.StringResult == "Value")
                            {
                                PromptStringOptions PrStrOpt = new PromptStringOptions("\nSpecify chainage: ");
                                PrStrOpt.AllowSpaces = false;
                                PromptResult PrStrRes = ed.GetString(PrStrOpt);
                                textKm = PrStrRes.StringResult;
                            }
                        }

                        double valKm = ExtractorNumar(textKm);
                        ed.WriteMessage("KM: {0}", valKm);

                        //Se verifica daca mai exista km respectiv in lista si se cere confirmare de suprascriere
                        if (listaKm.Contains(valKm))
                        {
                            PromptKeywordOptions PrKeyOpt = new PromptKeywordOptions("\nThe data for that chainage already exists! Overwrite?");
                            PrKeyOpt.Keywords.Add("Yes");
                            PrKeyOpt.Keywords.Add("No");
                            PrKeyOpt.AllowNone = false;
                            PrKeyOpt.AppendKeywordsToMessage = true;
                            PrKeyOpt.Keywords.Default = "Yes";
                            if (ed.GetKeywords(PrKeyOpt).StringResult == "No") continue;
                            else
                            {
                                listaKm.Remove(valKm);
                                String3D deSters = new String3D();
                                foreach (Punct3D p in listaCS)
                                {
                                    if (p.KM == valKm) deSters.Add(p);
                                }
                                foreach (Punct3D p in deSters) listaCS.Remove(p);
                            }
                        }


                        //Se face verificarea daca sistemul de coordonate curent este coplanar cu cel global
                        //In acest caz:
                        //Se stocheaza sistemul de coordonate curent si se trece in cel global
                        //UcsTable UCStable = (UcsTable)trans.GetObject(db.UcsTableId, OpenMode.ForRead);
                        //VA URMA
                        //In caz contrar se avertizeaza utilizatorul si se continua inregistrarea in sistemul de coordonate curent
                        //DE STUDIAT INTREGISTRAREA IN FORMAT TEXT A DEFINIRII SISTEMULUI DE COORDONATE;


                        //Se cere punctul de referinta
                        Point3d refPoint = new Point3d();
                        refPoint = ed.GetPoint("\nSelect the reference point for km " + valKm.ToString() + ": ").Value;

                        //Se cere cota punctului de referinta
                        string textCref = string.Empty;
                        PromptEntityOptions PrEntCrefOpt = new PromptEntityOptions("");
                        PrEntCrefOpt.Message = "\nIndicate Reference Point's Level (Text or Mtext)";
                        PrEntCrefOpt.Keywords.Add("Value");
                        PrEntCrefOpt.Keywords.Add("eXit");
                        PrEntCrefOpt.AppendKeywordsToMessage = true;
                        PrEntCrefOpt.SetRejectMessage("\nThe selected object is not of required type! ");
                        PrEntCrefOpt.AddAllowedClass(typeof(DBText), true);
                        PrEntCrefOpt.AddAllowedClass(typeof(MText), true);
                        PrEntCrefOpt.AllowNone = true;
                        PromptEntityResult PrEntCrefRes = ed.GetEntity(PrEntCrefOpt);
                        if (PrEntCrefRes.Status == PromptStatus.OK)
                        {
                            ObjectId objID = PrEntCrefRes.ObjectId;
                            Autodesk.AutoCAD.DatabaseServices.DBObject obj = trans.GetObject(objID, OpenMode.ForRead);
                            if (obj is DBText)
                            {
                                DBText text = (DBText)obj;
                                textCref = text.TextString;
                            }
                            if (obj is MText)
                            {
                                MText mtext = (MText)obj;
                                textCref = mtext.Contents;
                            }
                        }
                        else if (PrEntCrefRes.Status == PromptStatus.Cancel)
                        {
                            ed.WriteMessage("\nCommand aborted. Exiting!");
                            goto final;
                        }
                        else if (PrEntCrefRes.StringResult != string.Empty)
                        {
                            if (PrEntCrefRes.StringResult == "eXit")
                            {
                                ed.WriteMessage("\nCommand aborted. Exiting!");
                                goto final;
                            }
                            else if (PrEntCrefRes.StringResult == "Value")
                            {
                                PromptStringOptions PrStrOpt = new PromptStringOptions("\nSpecify the level of the reference point: ");
                                PrStrOpt.AllowSpaces = false;
                                PromptResult PrStrRes = ed.GetString(PrStrOpt);
                                textCref = PrStrRes.StringResult;
                            }
                        }

                        double valCref = ExtractorNumar(textCref);
                        ed.WriteMessage("Reference point level: {0}", valCref);

                        //Se cere offsetul punctului de referinta
                        string textOff = string.Empty;
                        double valOff = -999;
                        do
                        {
                            List<string> permise = new List<string>() { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "." };
                            PromptStringOptions PrStrOpt = new PromptStringOptions("\nSpecify the offset of the reference point");
                            PrStrOpt.DefaultValue = 0.ToString("F3");
                            PrStrOpt.UseDefaultValue = true;
                            string textInit = ed.GetString(PrStrOpt).StringResult;
                            foreach (char c in textInit.ToCharArray())
                            {
                                if (textOff == string.Empty && c.ToString() == "-") textOff = "-";
                                else if (permise.Contains(c.ToString())) textOff = textOff + c;
                            }
                            double.TryParse(textOff, out valOff);
                            if (valOff == -999)
                            {
                                ed.WriteMessage("\nOffset input format incorrect! ");
                                textOff = string.Empty;
                            }
                            else ed.WriteMessage("Reference point offset: {0}", valOff);
                        } while (valOff == -999);

                        //Se calculeaza coordonatele originii sistemului de coordonate al sectiunii
                        //si se inregistreaza sectiunea in lista
                        Punct3D CS = new Punct3D();
                        CS.KM = valKm;
                        CS.Offset = refPoint.X - valOff;
                        CS.Z = refPoint.Y - valCref;
                        listaKm.Add(valKm);
                        listaCS.Add(CS);
                        ed.WriteMessage("\nCoordinate system for recorded succesfully ({0}). Recording next Chainage. ",
                            CS.toString(Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 3, true));
                    }
                //Se completeaza fisierul cu sisteme de coordonate si se iese din comanda
                final:
                    listaCS.Sort(
                        delegate(Punct3D p1, Punct3D p2)
                        {
                            return p1.KM.CompareTo(p2.KM);
                        }
                    );
                    listaCS.ExportPoints(fisier.FullName, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 3, false);
                    ed.WriteMessage("\nCoordintates Systems file updated successfuly!");

                }
                #endregion
            }

            //Comanda pentru citirea sectiunilor si scrierea lor in fisiere *.csv
            [CommandMethod("ECS")]
            public void ecs()
            {
                Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = acadDoc.Editor;
                Database db = HostApplicationServices.WorkingDatabase;

                //Se trece la ucs global
                ed.CurrentUserCoordinateSystem = new Matrix3d(new double[16]{
    1.0, 0.0, 0.0, 0.0,
    0.0, 1.0, 0.0, 0.0,
    0.0, 0.0, 1.0, 0.0,
    0.0, 0.0, 0.0, 1.0});

                string caleDocAcad = acadDoc.Database.Filename;
                //Daca fisierul este unul nou, ne salvat inca, calea apartine sablonului folosit pentru acesta.
                //Se verifica daca fisierul este unul sablon, se atentioneaza utilizatorul si se paraseste programul.
                if (caleDocAcad.EndsWith(".dwt") == true)
                {
                    ed.WriteMessage("\nThe current drawing is a template file (*.dwt). Exiting program! ");
                    return;
                }
                caleDocAcad = HostApplicationServices.Current.FindFile(acadDoc.Name, acadDoc.Database, FindFileHint.Default);

                //Cautarea fisierului cu sisteme de referinta ale sectiunilor
                string calefisierSCS = caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + ".SCS";
                FileInfo fisierSCS = new FileInfo(calefisierSCS);
                if (fisierSCS.Exists == false)
                {
                    PromptKeywordOptions PrKeyOpt = new PromptKeywordOptions("\nThe coordinate system file does not exist! Run SCS command?");
                    PrKeyOpt.Keywords.Add("Yes");
                    PrKeyOpt.Keywords.Add("No");
                    PrKeyOpt.Keywords.Default = "Yes";
                    PrKeyOpt.AppendKeywordsToMessage = true;
                    PrKeyOpt.AllowNone = false;
                    if (ed.GetKeywords(PrKeyOpt).StringResult == "Yes") acadDoc.SendStringToExecute("scs", true, true, false);
                    else return;
                }

                //Citirea fisierului cu sisteme de referinta ale sectiunilor
                String3D listaCS = new String3D();
                listaCS.ImportPoints(fisierSCS.FullName, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma);
                listaCS.Sort(
                        delegate(Punct3D p1, Punct3D p2)
                        {
                            return p1.KM.CompareTo(p2.KM);
                        }
                    );
                ed.WriteMessage("\n{0} chainages read from km {1} to km {2}", listaCS.Count, listaCS[0].KM, listaCS[listaCS.Count - 1].KM);
                //ed.WriteMessage("\n{0} chainages read ", listaCS.Count);
                List<double> listaKm = new List<double>();
                foreach (Punct3D p in listaCS) listaKm.Add(p.KM);

                //Se cere descrierea sectiunilor si se creeaza fisierul txt
                string descriere = ed.GetString("\nSpecify the description of the sections to be read: ").StringResult;
                string zi = DateTime.Now.Date.Day.ToString();
                if (DateTime.Now.Date.Day < 10) zi = 0 + zi;
                string luna = DateTime.Now.Date.Month.ToString();
                if (DateTime.Now.Date.Month < 10) luna = 0 + luna;
                string an = DateTime.Now.Date.Year.ToString();
                string caleFisierRezultat = caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + "-" + descriere + "-"
                    + zi + luna + an + ".txt";
                if (new FileInfo(caleFisierRezultat).Exists == false)
                {
                    StreamWriter scriitor = new StreamWriter(caleFisierRezultat);
                    scriitor.Dispose();
                }
                else
                {
                    StreamWriter scriitor = new StreamWriter(caleFisierRezultat, true);
                    scriitor.WriteLine("File appended on {0} at {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString());
                    scriitor.Dispose();
                }

                //Bucla Principala - sugereaza km si cere polilinia oferind si optiuni suplimentare
                bool gata = false;
                int nrCS = 0;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    while (gata == false)
                    {
                        Punct3D CS = listaCS[nrCS];
                        String3D sectiune = new String3D();

                        //Optiuni selectie cu cuvinte cheie
                        PromptSelectionOptions PrSelPolyOpt = new PromptSelectionOptions();
                        PrSelPolyOpt.Keywords.Add("Km");
                        PrSelPolyOpt.Keywords.Add("eXit");
                        string PrSelPolyKws = PrSelPolyOpt.Keywords.GetDisplayString(true);
                        PrSelPolyOpt.MessageForAdding = "\nSelect the polyline(s) for km " + CS.KM.ToString() + " " + PrSelPolyKws;
                        PrSelPolyOpt.SinglePickInSpace = true;
                        bool cont = false;
                        bool retur = false;

                        //Metoda care se apeleaza la selectia unui cuvant cheie
                        PrSelPolyOpt.KeywordInput += delegate(object sender, SelectionTextInputEventArgs e)
                        {
                            switch (e.Input)
                            {
                                case "Km":
                                    double valKm = -999;
                                    PromptDoubleOptions PrDblOpt = new PromptDoubleOptions("\nSpecify the desired chainage value: ");
                                    PrDblOpt.Keywords.Add("List");
                                    PrDblOpt.Keywords.Add("eXit");
                                    PrDblOpt.AppendKeywordsToMessage = true;
                                    PromptDoubleResult PrDblRes = ed.GetDouble(PrDblOpt);
                                    valKm = PrDblRes.Value;
                                    if (PrDblRes.Status == PromptStatus.OK && valKm != -999)
                                    {
                                        if (listaKm.Contains(valKm)) nrCS = listaKm.IndexOf(valKm);
                                        else ed.WriteMessage("\nThe chainage is not defined in the section coordinate systems file! ");
                                    }
                                    else if (PrDblRes.StringResult == "List")
                                    {
                                        foreach (Punct3D p in listaCS)
                                        {
                                            string textKm = p.toString(Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 3, true);
                                            ed.WriteMessage("\nKm: {0}", textKm.Remove(textKm.IndexOf(',')));
                                        }
                                    }
                                    else cont = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    //acadDoc.SendStringToExecute(((char)27).ToString(), true, true, true); //Cancel character
                                    break;
                                case "eXit":
                                    //ed.WriteMessage("\nExiting program!");
                                    //acadDoc.SendStringToExecute(((char)27).ToString(), true, true, true); //Cancel character
                                    retur = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    break;
                                default:
                                    cont = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    break;
                            }
                        };



                        //Filtru selectie
                        TypedValue[] tvs = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                    };
                        SelectionFilter sf = new SelectionFilter(tvs);

                        PromptSelectionResult PrSelPolyRes = ed.GetSelection(PrSelPolyOpt, sf);

                        //Se actioneaza conform selectiei cuvantului cheie
                        if (cont) continue;
                        if (retur)
                        {
                            ed.WriteMessage("\nExiting program!");
                            return;
                        }

                        //Se citeste polilinia selectata si se completeaza fisierul rezultat
                        if (PrSelPolyRes.Status == PromptStatus.OK)
                        {
                            List<Point3d> listavertecsi = new List<Point3d>();
                            ObjectId[] objIds = PrSelPolyRes.Value.GetObjectIds();
                            foreach (ObjectId polyId in objIds)
                            {
                                Polyline poly = (Polyline)trans.GetObject(polyId, OpenMode.ForRead);

                                for (int i = 0; i < poly.NumberOfVertices; i++)
                                {
                                    listavertecsi.Add(poly.GetPoint3dAt(i));
                                }
                            }
                            List<Point3d> listatrans = Transfom2Param(listavertecsi, CS.Offset, CS.Z, 0, 0);
                            StreamWriter scriitor2 = new StreamWriter(caleFisierRezultat, true); //Vezi comentariile de mai jos
                            foreach (Point3d p in listatrans)
                            {
                                Punct3D p3d = new Punct3D();
                                p3d.KM = CS.KM;
                                p3d.Offset = p.X;
                                p3d.Z = p.Y;
                                p3d.D = descriere;
                                sectiune.Add(p3d); //Devine inutil fara modificarea de mai jos dar lasam asa pt. moment
                                scriitor2.WriteLine(p3d.KM.ToString("F4") + "," + p3d.Offset.ToString("F4") + ","
                                    + p3d.Z.ToString("F4") + "," + p3d.D);
                            }
                            //Utilitarul StrinUtil trebuie modificat pentru a elimina suprascrierea automata a fisierelor in care se exporta puncte
                            //sectiune.ExportPoints(caleFisierRezultat, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 3, false);
                            scriitor2.Dispose();
                            ed.WriteMessage("\nSection export file updated successfully.");
                            if (nrCS == listaCS.Count - 1)
                            {
                                ed.WriteMessage("\nEnd of section coordinate systems file reached! Exiting Program.");
                                return;
                            }
                            else nrCS = nrCS + 1;
                        }
                        //else
                        //{
                        //    //if (cont) continue;
                        //    if (retur) return;
                        //}
                    }
                }

            }

            //Comanda pentru desenarea sectiunilor dintr-un fisier text, cu ajutorul fisierului cu sisteme de referinta ale sectiunilor
            [CommandMethod("DCS")]
            public void dcs()
            {
                Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = acadDoc.Editor;
                Database db = HostApplicationServices.WorkingDatabase;

                string caleDocAcad = acadDoc.Database.Filename;
                //Daca fisierul este unul nou, ne salvat inca, calea apartine sablonului folosit pentru acesta.
                //Se verifica daca fisierul este unul sablon, se atentioneaza utilizatorul si se paraseste programul.
                if (caleDocAcad.EndsWith(".dwt") == true)
                {
                    ed.WriteMessage("\nThe current drawing is a template file (*.dwt). Exiting program! ");
                    return;
                }
                caleDocAcad = HostApplicationServices.Current.FindFile(acadDoc.Name, acadDoc.Database, FindFileHint.Default);

                //Cautarea fisierului cu sisteme de referinta ale sectiunilor
                string calefisierSCS = caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + ".SCS";
                FileInfo fisierSCS = new FileInfo(calefisierSCS);
                if (fisierSCS.Exists == false)
                {
                    PromptKeywordOptions PrKeyOpt = new PromptKeywordOptions("\nThe coordinate system file does not exist! Run SCS command?");
                    PrKeyOpt.Keywords.Add("Yes");
                    PrKeyOpt.Keywords.Add("No");
                    PrKeyOpt.Keywords.Default = "Yes";
                    PrKeyOpt.AppendKeywordsToMessage = true;
                    PrKeyOpt.AllowNone = false;
                    if (ed.GetKeywords(PrKeyOpt).StringResult == "Yes") acadDoc.SendStringToExecute("scs", true, false, true);
                    else return;
                }

                //Citirea fisierului cu sisteme de referinta ale sectiunilor
                String3D listaCS = new String3D();
                listaCS.ImportPoints(fisierSCS.FullName, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma);
                listaCS.Sort(
                        delegate(Punct3D p1, Punct3D p2)
                        {
                            return p1.KM.CompareTo(p2.KM);
                        }
                    );
                ed.WriteMessage("\n{0} chainages read from km {1} to km {2}", listaCS.Count, listaCS[0].KM, listaCS[listaCS.Count - 1].KM);
                //ed.WriteMessage("\n{0} chainages read ", listaCS.Count);
                List<double> listaKm = new List<double>();
                foreach (Punct3D p in listaCS) listaKm.Add(p.KM);

                //Selectia fisierului .txt cu sectiuni
                OpenFileDialog fileDia = new OpenFileDialog("Section file selection", null, "txt",
                    "File Selection", OpenFileDialog.OpenFileDialogFlags.NoUrls);
                System.Windows.Forms.DialogResult dr = fileDia.ShowDialog();
                if (dr != System.Windows.Forms.DialogResult.OK)
                {
                    ed.WriteMessage("\n Incorrect selection of section file! Aborting.");
                    return;
                }
                string caleSectiuni = fileDia.Filename;


                //Citirea fisierului .txt cu sectiuni
                String3D Puncte = new String3D();
                Puncte.ImportPoints(caleSectiuni, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma);
                if (Puncte.Count == 0)
                {
                    ed.WriteMessage("\nNo points could be read from the section file! Aborting.");
                    return;
                }
                else ed.WriteMessage("\n{0} points were read.");

                //Impartirea punctelor pe sectiuni
                List<double> KmSectiuni = new List<double>();
                List<String3D> Sectiuni = new List<String3D>();
                foreach (Punct3D Punct in Puncte)
                {
                    if (KmSectiuni.Count == 0 || Punct.KM != KmSectiuni[KmSectiuni.Count - 1])
                    {
                        KmSectiuni.Add(Punct.KM);
                        String3D Sectiune = new String3D();
                        Sectiune.Add(Punct);
                        Sectiuni.Add(Sectiune);
                    }
                    else
                    {
                        Sectiuni[Sectiuni.Count - 1].Add(Punct);
                    }
                }
                ed.WriteMessage("\nSections were found for the following chainages: ");
                foreach (double Km in KmSectiuni) ed.WriteMessage("\nKm: {0}", Km);

                //Se trece la ucs global
                ed.CurrentUserCoordinateSystem = new Matrix3d(new double[16]{
    1.0, 0.0, 0.0, 0.0,
    0.0, 1.0, 0.0, 0.0,
    0.0, 0.0, 1.0, 0.0,
    0.0, 0.0, 0.0, 1.0});

                //Desenarea sectiunilor
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    foreach (String3D sect in Sectiuni)
                    {
                        if (listaKm.Contains(sect[0].KM) == false)
                        {
                            ed.WriteMessage("\nNo coordinate system found for chainage: {0}! Section skipped.", sect[0].KM);
                            continue;
                        }
                        else if (sect.Count < 2)
                        {
                            ed.WriteMessage("\nThe section at chainage: {0} has too few points to be drawn! Section skipped.", sect[0].KM);
                            continue;
                        }
                        Polyline poly = new Polyline();
                        foreach (Punct3D pct in sect)
                        {
                            double deltaOffset = listaCS.Find(CS => CS.KM == pct.KM).Offset;
                            double deltaZ = listaCS.Find(CS => CS.KM == pct.KM).Z;
                            Point2d p = new Point2d(pct.Offset + deltaOffset, pct.Z + deltaZ);
                            poly.AddVertexAt(poly.NumberOfVertices, p, 0, 0, 0);
                        }
                        btr.AppendEntity(poly);
                        trans.AddNewlyCreatedDBObject(poly, true);
                        //ed.WriteMessage("\nThe section at chainage: {0} was drawn successfully!", sect[0].KM);
                    }
                    trans.Commit();
                }

            }

            //Comanda pentru exportarea blocurilor cu sectiuni transversale intr-un alt desen
            [CommandMethod("EBS")]
            public void ebs()
            {
                Document acadDoc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = acadDoc.Editor;
                Database db = HostApplicationServices.WorkingDatabase;

                //Se trece la ucs global
                ed.CurrentUserCoordinateSystem = new Matrix3d(new double[16]{
    1.0, 0.0, 0.0, 0.0,
    0.0, 1.0, 0.0, 0.0,
    0.0, 0.0, 1.0, 0.0,
    0.0, 0.0, 0.0, 1.0});

                string caleDocAcad = acadDoc.Database.Filename;
                //Daca fisierul este unul nou, ne salvat inca, calea apartine sablonului folosit pentru acesta.
                //Se verifica daca fisierul este unul sablon, se atentioneaza utilizatorul si se paraseste programul.
                if (caleDocAcad.EndsWith(".dwt") == true)
                {
                    ed.WriteMessage("\nThe current drawing is a template file (*.dwt). Exiting program! ");
                    return;
                }
                caleDocAcad = HostApplicationServices.Current.FindFile(acadDoc.Name, acadDoc.Database, FindFileHint.Default);

                //Cautarea fisierului cu sisteme de referinta ale sectiunilor
                string calefisierSCS = caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + ".SCS";
                FileInfo fisierSCS = new FileInfo(calefisierSCS);
                if (fisierSCS.Exists == false)
                {
                    PromptKeywordOptions PrKeyOpt = new PromptKeywordOptions("\nThe coordinate system file does not exist! Run SCS command?");
                    PrKeyOpt.Keywords.Add("Yes");
                    PrKeyOpt.Keywords.Add("No");
                    PrKeyOpt.Keywords.Default = "Yes";
                    PrKeyOpt.AppendKeywordsToMessage = true;
                    PrKeyOpt.AllowNone = false;
                    if (ed.GetKeywords(PrKeyOpt).StringResult == "Yes") acadDoc.SendStringToExecute("scs", true, true, false);
                    else return;
                }

                //Citirea fisierului cu sisteme de referinta ale sectiunilor
                String3D listaCS = new String3D();
                listaCS.ImportPoints(fisierSCS.FullName, Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma);
                listaCS.Sort(
                        delegate(Punct3D p1, Punct3D p2)
                        {
                            return p1.KM.CompareTo(p2.KM);
                        }
                    );
                ed.WriteMessage("\n{0} chainages read from km {1} to km {2}", listaCS.Count, listaCS[0].KM, listaCS[listaCS.Count - 1].KM);
                //ed.WriteMessage("\n{0} chainages read ", listaCS.Count);
                List<double> listaKm = new List<double>();
                foreach (Punct3D p in listaCS) listaKm.Add(p.KM);

                //Crearea bazei de date si a desenului tinta
                string caletinta = caleDocAcad.Remove(caleDocAcad.LastIndexOf('.')) + "-EBS" + ".dwg";
                //Document doctinta = Application.DocumentManager.
                Database newDb = new Database();
                using (Transaction tr = newDb.TransactionManager.StartTransaction())
                {
                    BlockTableRecord newBT = new BlockTableRecord();
              

                //Bucla Principala - sugereaza km si cere obiectele blocului oferind si optiuni suplimentare
                List<SelectionSet> listaselectii = new List<SelectionSet>();
                bool gata = false;
                int nrCS = 0;

                while (gata == false)
                {
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        Punct3D CS = listaCS[nrCS];
                        String3D sectiune = new String3D();
                        BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);

                        //Optiuni selectie cu cuvinte cheie
                        PromptSelectionOptions PrSelOpt = new PromptSelectionOptions();
                        PrSelOpt.Keywords.Add("Km");
                        PrSelOpt.Keywords.Add("eXit");
                        string PrSelPolyKws = PrSelOpt.Keywords.GetDisplayString(true);
                        PrSelOpt.MessageForAdding = "\nSelect the opbjects for km " + CS.KM.ToString() + " " + PrSelPolyKws;
                        PrSelOpt.SinglePickInSpace = false;
                        bool cont = false;
                        bool retur = false;

                        //Metoda care se apeleaza la selectia unui cuvant cheie
                        PrSelOpt.KeywordInput += delegate(object sender, SelectionTextInputEventArgs e)
                        {
                            switch (e.Input)
                            {
                                case "Km":
                                    double valKm = -999;
                                    PromptDoubleOptions PrDblOpt = new PromptDoubleOptions("\nSpecify the desired chainage value: ");
                                    PrDblOpt.Keywords.Add("List");
                                    PrDblOpt.Keywords.Add("eXit");
                                    PrDblOpt.AppendKeywordsToMessage = true;
                                    PromptDoubleResult PrDblRes = ed.GetDouble(PrDblOpt);
                                    valKm = PrDblRes.Value;
                                    if (PrDblRes.Status == PromptStatus.OK && valKm != -999)
                                    {
                                        if (listaKm.Contains(valKm)) nrCS = listaKm.IndexOf(valKm);
                                        else ed.WriteMessage("\nThe chainage is not defined in the section coordinate systems file! ");
                                    }
                                    else if (PrDblRes.StringResult == "List")
                                    {
                                        foreach (Punct3D p in listaCS)
                                        {
                                            string textKm = p.toString(Punct3D.Format.KmOZ, Punct3D.DelimitedBy.Comma, 3, true);
                                            ed.WriteMessage("\nKm: {0}", textKm.Remove(textKm.IndexOf(',')));
                                        }
                                    }
                                    else cont = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    //acadDoc.SendStringToExecute(((char)27).ToString(), true, true, true); //Cancel character
                                    break;
                                case "eXit":
                                    //ed.WriteMessage("\nExiting program!");
                                    //acadDoc.SendStringToExecute(((char)27).ToString(), true, true, true); //Cancel character
                                    retur = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    break;
                                default:
                                    cont = true;
                                    acadDoc.SendStringToExecute(((char)32).ToString(), true, false, false); //Space character
                                    break;
                            }
                        };



                        //Selectia obiectelor blocului si crearea lui;
                        //PromptSelectionOptions PrSelOpt = new PromptSelectionOptions();
                        //PrSelOpt.MessageForAdding = "\nSelect the objects for block ok km " + CS.KM.ToString() + ": ";
                        PromptSelectionResult PrSelObj = ed.GetSelection(PrSelOpt);
                        if (PrSelObj.Status == PromptStatus.OK)
                        {
                            ObjectIdCollection PrSelObjIdColl = new ObjectIdCollection(PrSelObj.Value.GetObjectIds());

                            if (PrSelObjIdColl.Count > 0)
                            {
                                string newBlockName = "XS_Km" + CS.KM.ToString();
                                if (bt.Has(newBlockName))
                                {
                                    if (ed.GetKeywords("\nSection block already exists! Overwrite?", "Yes", "No").StringResult == "No")
                                    {
                                        continue;
                                    }
                                }
                                BlockTableRecord newBlock = new BlockTableRecord();
                                newBlock.Name = newBlockName;
                                newBlock.Origin = new Point3d(CS.toArray(Punct3D.Format.ENZ));
                                ObjectId newBlockId = bt.Add(newBlock);
                                trans.AddNewlyCreatedDBObject(newBlock, true);
                                IdMapping mapping = new IdMapping();
                                db.DeepCloneObjects(PrSelObjIdColl, newBlockId, mapping, false);
                                nrCS++;

                            }
                            else
                            {
                                ed.WriteMessage("\nNo objects were selected!");
                                cont = true;
                            }
                        }

                        //Se actioneaza conform selectiei cuvantului cheie
                        if (cont) continue;
                        if (retur)
                        {
                            ed.WriteMessage("\nExiting program!");
                            return;
                        }
                        trans.Commit();
                    }

                }
            }
                }

            //Comanda pentru importarea blocurilor cu sectiuni transversale dintr-un alt desen
            [CommandMethod("IBS")]
            public void ibs()
            {
            }

            public static Ovidiu.x64.General.Configurator GetConfig()
            {
                //Cauta fisierul de configurare in calea implicita
                string cale;
                cale = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Optiuni.cfg";
                cale = Path.GetFullPath(cale);
                if (File.Exists(cale) == false)
                {
                    try
                    {
                        using (StreamWriter scriitor = new StreamWriter(cale))
                        {
                            scriitor.Write("\r\n");
                        }
                    }
                    catch
                    {
                    }
                }
                return new Ovidiu.x64.General.Configurator(cale);
            }

            //Functie care extrage primul numar intalnit intr-un element text.
            //Recunoaste separatorul decimal si ignora caracterul '+'.
            public static double ExtractorNumar(string txt)
            {
                double rezultat = -999.999;
                txt = txt.Replace("+", "");
                List<string> permise = new List<string>() { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "." };
                int start = -999;
                int end = -999;
                for (int i = 0; i < txt.Length; i++)
                {
                    string c = txt.Substring(i, 1);
                    if (start == -999 && permise.Contains(c))
                        start = i;
                    if (end == -999 && start != -999 && permise.Contains(c) == false)
                        end = i;
                }
                if (start == -999) start = 0;
                if (end == -999) end = txt.Length;
                double.TryParse(txt.Substring(start, end - start), out rezultat);
                return rezultat;
            }

            public static List<Point3d> Transfom2Param(List<Point3d> puncte, double oldX, double oldY, double newX, double newY)
            {
                List<Point3d> rezultat = new List<Point3d>();

                double deltaX = newX - oldX;
                double deltaY = newY - oldY;

                foreach (Point3d punct in puncte)
                {
                    rezultat.Add(new Point3d(punct.X + deltaX, punct.Y + deltaY, punct.Z));
                }

                return rezultat;
            }
        }

}
