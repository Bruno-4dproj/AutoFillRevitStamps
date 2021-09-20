using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.IO;
using OfficeOpenXml;
using System.Collections;
using System.Data.OleDb;
using Autodesk.Revit.ApplicationServices;

namespace RevitAddin
{

    [Transaction(TransactionMode.Manual)]
    public class CreateSheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("4A6AC663-2FF3-409C-9C00-24390C773550"));
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {

            #region Obrigatorio do IExternarCommand, Comando para acessar o arquivo e os comandos no revit
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            #endregion

            
            Form_Carimbo fc = new Form_Carimbo();
            fc.ShowDialog();


            #region Verifica qual Disciplina foi escolhida
            if (MinhasVariaveis.stats == "Cancelado")
            {
                return Result.Cancelled;
            }
            if (MinhasVariaveis.stats == "01")
            {
                MinhasVariaveis.disciplina = "Select * from [MEC$]";
            }
            if (MinhasVariaveis.stats == "02")
            {
                MinhasVariaveis.disciplina = "Select * from [PLU$]";
            }
            if (MinhasVariaveis.stats == "03")
            {
                MinhasVariaveis.disciplina = "Select * from [ELE$]";
            }
            if (MinhasVariaveis.stats == "04")
            {
                MinhasVariaveis.disciplina = "Select * from [LVS$]";
            }
            #endregion

            //Verificação caso o Usuario não carrege o arquivo, não executar os comando sem necessidade, pois vai dar erro.
            if (MinhasVariaveis.path != null)
            {
                List<Entidades> verificaçãoExcel = ObterCamposExcel(MinhasVariaveis.disciplina);
                if (verificaçãoExcel != null)
                {
                    using (Transaction t = new Transaction(doc, "CarimboRevit"))
                    {
                        t.Start();
                        //Metodo para criar uma  folha
                        CreateSheets(uidoc, doc);

                        //Inserir views na folha
                        PlaceView(uidoc, doc);
                        
                        //preenche o carimbo
                        PreencherCarimbo(uidoc, doc);

                        //Preencher parametros de projeto
                        PreencherParametros(uidoc, doc);

                        //Ativa o campo de revisão da folh
                        CheckRevisão(uidoc, doc);

                        t.Commit();
                    }
                }
                

            }
            else
            {
                return Result.Cancelled;
            }

            //Obrigatorio do IexternalCommand, deve sempre retornar o exito ou falha.
            return Result.Succeeded;

        }

        #region Comando para Criar as folhas
        /// <summary>
        /// Comando para criar as folhas no Revit
        /// </summary>
        /// <param name="uidoc">Variavel usada pelo revit</param>
        /// <param name="doc">Variavel usada pelo revit</param>
        public void CreateSheets(UIDocument uidoc, Document doc)
        {

            //Get Family Symbol com o nome do simbolo que eu quero usar.
            FamilySymbol tBlock = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .FirstOrDefault<FamilySymbol>(o => o.Name == "JMBF_A0_padrão");


            
            //Cria uma lista com os dados da planilha excel
            List<Entidades> listaFolha = ObterCamposExcel(MinhasVariaveis.disciplina);
            if (listaFolha != null)
            {

                foreach (var a in listaFolha)
                {

                    //Filtrar as folhas no documento Revit
                    ViewSheet viewSheet = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Sheets)
                        .Cast<ViewSheet>()
                        .FirstOrDefault(s => s.SheetNumber == a.FOLHA);

                    try
                    {
                        //using (Transaction t = new Transaction(doc))
                        //{
                        //t.Start();

                        if (viewSheet == null)
                        {
                            if(a.FOLHA != "")
                            {
                                //Criando a folha
                                ViewSheet vSheet = ViewSheet.Create(doc, tBlock.Id);
                                vSheet.Name = a.NOMEFOLHA;
                                vSheet.SheetNumber = a.FOLHA;
                            }
                        }

                        //t.Commit();

                        //}
                    }
                    catch
                    {
                        TaskDialog.Show("Erro", "Não possivel criar a folha");
                    }
                }
            }
        }
        #endregion 

        #region Comando para inserir as Views 
        private void PlaceView(UIDocument uidoc, Document doc)
        {
            
            List<Entidades> listaFolha = ObterCamposExcel(MinhasVariaveis.disciplina);
            if (listaFolha != null)
            {

                List<String> erroView = new List<string>();

                foreach (var a in listaFolha)
                {


                    //Filtrar as Viwes Plans no revit pelo Nome da Folha
                    Element vPlan = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Views)
                        .FirstOrDefault(v => v.Name == a.NOMEFOLHA);

                    //Filtrar as folhas que existem no Revit
                    ViewSheet viewSheet = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Sheets)
                        .Cast<ViewSheet>()
                        .FirstOrDefault(s => s.SheetNumber == a.FOLHA);


                    if (vPlan != null)
                    {
                        if (viewSheet.ViewName != vPlan.ToString())
                        {
                            try
                            {
                                //using (Transaction t = new Transaction(doc))
                                //{
                                //Criar o ponto de inserção na folha
                                BoundingBoxUV outline = viewSheet.Outline;
                                double xu = (outline.Max.U + outline.Min.U) / 2;
                                double yv = (outline.Max.V + outline.Min.V) / 2;
                                XYZ midPoint = new XYZ(xu, yv, 0);

                                if (vPlan != null)
                                {
                                    //t.Start();

                                    //Inserir a view na planta
                                    Viewport vPort = Viewport.Create(doc, viewSheet.Id, vPlan.Id, midPoint);

                                    //t.Commit();
                                }
                                //}
                            }
                            catch
                            {
                                
                            }

                            
                        }
                    }
                    
                }

            }

        }
        #endregion

        #region Preencher carimbo no Revit
        /// <summary>
        /// Metodo para preencher o carimbo no Revit
        /// </summary>
        /// <param name="uidoc"></param>
        /// <param name="doc"></param>
        public void PreencherCarimbo(UIDocument uidoc, Document doc)
        {
            
            List<Entidades> listaFolha = ObterCamposExcel(MinhasVariaveis.disciplina);

            
            if (listaFolha != null)
            {
                try
                {

                    foreach (var a in listaFolha)
                    {
                        //Filtrar as folhas no documento Revit
                        ViewSheet viewSheet = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Sheets)
                        .Cast<ViewSheet>()
                        .FirstOrDefault(s => s.SheetNumber == a.FOLHA);


                        ParameterSet pSet = viewSheet.Parameters;

                        if (viewSheet.SheetNumber == a.FOLHA)
                        {
                            
                            foreach (Parameter param in pSet)
                            {
                                //Parametros padrão da folha
                                if (param.Definition.Name.Equals("Sheet Name"))
                                {
                                    param.Set(a.NOMEFOLHA);
                                }
                                if (param.Definition.Name.Equals("Titulo 01"))
                                {
                                    param.Set(a.TITULO01);
                                }
                                if (param.Definition.Name.Equals("Titulo 02"))
                                {
                                    param.Set(a.TITULO02);
                                }
                                if (param.Definition.Name.Equals("Titulo 03"))
                                {
                                    param.Set(a.TITULO03);
                                }
                                if (param.Definition.Name.Equals("Titulo 04"))
                                {
                                    param.Set(a.TITULO04);
                                }
                                if (param.Definition.Name.Equals("ESCALA"))
                                {
                                    param.Set(a.ESCALA);
                                }
                                if (param.Definition.Name.Equals("REVISÃO"))
                                {
                                    param.Set(a.REVISÃO);
                                }
                                if (param.Definition.Name.Equals("DATA R00"))
                                {
                                    string data = a.DATAR00.Substring(0, 6) + 
                                                  a.DATAR00.Substring(8, 2);
                                    param.Set(data);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R00"))
                                {
                                    param.Set(a.DESCRIÇÃOR00);
                                }
                                if (param.Definition.Name.Equals("DESENHO R00"))
                                {
                                    param.Set(a.DESENHOR00);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R00"))
                                {
                                    param.Set(a.VERIFICADOR00);
                                }
                                #region REVISÃO R01
                                //parametros da revisão R01
                                if (param.Definition.Name.Equals("DATA R01"))
                                {
                                    param.Set(a.DATAR01);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R01"))
                                {
                                    param.Set(a.DESCRIÇÃOR01);
                                }
                                if (param.Definition.Name.Equals("DESENHO R01"))
                                {
                                    param.Set(a.DESENHOR01);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R01"))
                                {
                                    param.Set(a.VERIFICADOR01);
                                }
                                #endregion
                                #region REVISÃO R02
                                //parametros da revisão R02
                                if (param.Definition.Name.Equals("DATA R02"))
                                {
                                    param.Set(a.DATAR01);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R02"))
                                {
                                    param.Set(a.DESCRIÇÃOR02);
                                }
                                if (param.Definition.Name.Equals("DESENHO R02"))
                                {
                                    param.Set(a.DESENHOR02);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R02"))
                                {
                                    param.Set(a.VERIFICADOR02);
                                }
                                #endregion
                                #region REVISÃO R03
                                //parametros da revisão R03
                                if (param.Definition.Name.Equals("DATA R03"))
                                {
                                    param.Set(a.DATAR03);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R03"))
                                {
                                    param.Set(a.DESCRIÇÃOR03);
                                }
                                if (param.Definition.Name.Equals("DESENHO R03"))
                                {
                                    param.Set(a.DESENHOR03);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R03"))
                                {
                                    param.Set(a.VERIFICADOR03);
                                }
                                #endregion
                                #region REVISÃO R04
                                //parametros da revisão R04
                                if (param.Definition.Name.Equals("DATA R04"))
                                {
                                    param.Set(a.DATAR04);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R04"))
                                {
                                    param.Set(a.DESCRIÇÃOR04);
                                }
                                if (param.Definition.Name.Equals("DESENHO R04"))
                                {
                                    param.Set(a.DESENHOR04);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R04"))
                                {
                                    param.Set(a.VERIFICADOR04);
                                }
                                #endregion
                                #region REVISÃO R05
                                //parametros da revisão R05
                                if (param.Definition.Name.Equals("DATA R05"))
                                {
                                    param.Set(a.DATAR05);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R05"))
                                {
                                    param.Set(a.DESCRIÇÃOR05);
                                }
                                if (param.Definition.Name.Equals("DESENHO R05"))
                                {
                                    param.Set(a.DESENHOR05);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R05"))
                                {
                                    param.Set(a.VERIFICADOR05);
                                }
                                #endregion
                                #region REVISÃO R06
                                //parametros da revisão R06
                                if (param.Definition.Name.Equals("DATA R06"))
                                {
                                    param.Set(a.DATAR06);
                                }
                                if (param.Definition.Name.Equals("DESCRIÇÃO R06"))
                                {
                                    param.Set(a.DESCRIÇÃOR06);
                                }
                                if (param.Definition.Name.Equals("DESENHO R06"))
                                {
                                    param.Set(a.DESENHOR06);
                                }
                                if (param.Definition.Name.Equals("VERIFICADO R06"))
                                {
                                    param.Set(a.VERIFICADOR01);
                                }
                                #endregion
                            }
                                                        
                        }
                        else
                        {

                        }
                    }
                }
                catch
                {
                    //TaskDialog.Show("Catch", "Não foi possivel executar Preencher Carbimbo");
                }
            }
        }
        #endregion

        #region Comando para preencher Parametros de projeto
        public void PreencherParametros(UIDocument uidoc, Document doc)
        {
            List<ParametrosDeProjeto> listaParametrosProjeto = ObterParametrosProjeto();

            List<Parameter> listParamProj = new List<Parameter>();

            if(listaParametrosProjeto != null)
            {
                try
                {
                    foreach (var a in listaParametrosProjeto)
                    {
                        ProjectInfo paramProjeto = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_ProjectInformation)
                        .Cast<ProjectInfo>()
                        .FirstOrDefault();

                        ParameterSet paramSet = paramProjeto.Parameters;

                        foreach (Parameter par in paramSet)
                        {
                            if(par.Definition.Name.Equals("Disciplina"))
                            {
                                if (MinhasVariaveis.disciplina == "Select * from [MEC$]")
                                {
                                    par.Set("MEC");
                                }
                                if (MinhasVariaveis.disciplina == "Select * from [PLU$]")
                                {
                                    par.Set("PLU");
                                }
                                if (MinhasVariaveis.disciplina == "Select * from [ELE$]")
                                {
                                    par.Set("ELE");
                                }
                                if (MinhasVariaveis.disciplina == "Select * from [LVS$]")
                                {
                                    par.Set("LVS");
                                }
                            }
                            if (par.Definition.Name.Equals("PROJETO"))
                            {
                                par.Set(a.PROJETO);
                            }
                            if (par.Definition.Name.Equals("FASE PROJETO"))
                            {
                                par.Set(a.FASEPROJETO);
                            }
                            if (par.Definition.Name.Equals("ABREVIAÇÃO FASE PROJETO"))
                            {
                                par.Set(a.ABREVIAÇÃOFASE);
                            }
                            if (par.Definition.Name.Equals("ABREVIAÇÃO CLIENTE"))
                            {
                                par.Set(a.ABREVIAÇÃOCLIENTE);
                            }
                            if (par.Definition.Name.Equals("Project Number"))
                            {
                                par.Set(a.NUMEROPROJETO);
                            }
                            if (par.Definition.Name.Equals("Project Issue Date"))
                            {
                                par.Set(a.DATAINICIAL);
                            }
                            if (par.Definition.Name.Equals("Project Address"))
                            {
                                par.Set(a.ENDEREÇO);
                            }
                        }

                    }
                     
                        
                }
                catch
                {

                }
            }


        }

        #endregion

        #region Leitura da planilha Excel por OLEDB
        public static List<Entidades> ObterCamposExcel(string comandoSql)
        {
            MinhasVariaveis.connectionStringBase = @"provider = Microsoft.ACE.OLEDB.12.0;";
            MinhasVariaveis.dataSourceBase = "data source =";
            MinhasVariaveis.connectionPropertiesBase = ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;IMAX=1;'";


            string connectionString = MinhasVariaveis.connectionStringBase +
                                      MinhasVariaveis.dataSourceBase +
                                      MinhasVariaveis.path +
                                      MinhasVariaveis.connectionPropertiesBase;


            OleDbConnection connect = new OleDbConnection(connectionString);
            OleDbCommand comando = new OleDbCommand(comandoSql, connect);
            List<Entidades> listaCarimbo = new List<Entidades>();


            //tenta connectar com o excel.
            try
            {
                connect.Open();
                OleDbDataReader rd = comando.ExecuteReader();

                //depois de connectado ele vai ler as linhas enquanto houver informação.
                while (rd.Read())
                {
                    listaCarimbo.Add(new Entidades()
                    {
                        //Campos R00 e geral das folhas
                        FOLHA = rd["FOLHA"].ToString(),
                        NOMEFOLHA = rd["NOME FOLHA"].ToString(),
                        TITULO01 = rd["TITULO01"].ToString(),
                        TITULO02 = rd["TITULO02"].ToString(),
                        TITULO03 = rd["TITULO03"].ToString(),
                        TITULO04 = rd["TITULO04"].ToString(),
                        REVISÃO = rd["REVISÃO"].ToString(),
                        DATAR00 = rd["DATAR00"].ToString(),
                        DESCRIÇÃOR00 = rd["DESCRIÇÃOR00"].ToString(),
                        DESENHOR00 = rd["DESENHOR00"].ToString(),
                        VERIFICADOR00 = rd["VERIFICADOR00"].ToString(),
                        ESCALA = rd["ESCALA"].ToString(),

                        //campos para revisão 01
                        DATAR01 = rd["DATAR01"].ToString(),
                        DESCRIÇÃOR01 = rd["DESCRIÇÃOR01"].ToString(),
                        DESENHOR01 = rd["DESENHOR01"].ToString(),
                        VERIFICADOR01 = rd["VERIFICADOR01"].ToString(),

                        //campos para revisão 02
                        DATAR02 = rd["DATAR02"].ToString(),
                        DESCRIÇÃOR02 = rd["DESCRIÇÃOR02"].ToString(),
                        DESENHOR02 = rd["DESENHOR02"].ToString(),
                        VERIFICADOR02 = rd["VERIFICADOR02"].ToString(),

                        //campos para revisão 03
                        DATAR03 = rd["DATAR03"].ToString(),
                        DESCRIÇÃOR03 = rd["DESCRIÇÃOR03"].ToString(),
                        DESENHOR03 = rd["DESENHOR03"].ToString(),
                        VERIFICADOR03 = rd["VERIFICADOR03"].ToString(),

                        //campos para revisão 04
                        DATAR04 = rd["DATAR04"].ToString(),
                        DESCRIÇÃOR04 = rd["DESCRIÇÃOR04"].ToString(),
                        DESENHOR04 = rd["DESENHOR04"].ToString(),
                        VERIFICADOR04 = rd["VERIFICADOR04"].ToString(),

                        //campos para revisão 05
                        DATAR05 = rd["DATAR05"].ToString(),
                        DESCRIÇÃOR05 = rd["DESCRIÇÃOR05"].ToString(),
                        DESENHOR05 = rd["DESENHOR05"].ToString(),
                        VERIFICADOR05 = rd["VERIFICADOR05"].ToString(),

                        //campos para revisão 06
                        DATAR06 = rd["DATAR06"].ToString(),
                        DESCRIÇÃOR06 = rd["DESCRIÇÃOR06"].ToString(),
                        DESENHOR06 = rd["DESENHOR06"].ToString(),
                        VERIFICADOR06 = rd["VERIFICADOR06"].ToString(),

                    });

                }

                //Se a lista tem algum valor preenchido, ela retorna "listaCarimbo".
                if (listaCarimbo.Count() > 0)
                    return listaCarimbo;
                //Se a lista esta vazia, não retorna nada.
                else
                    return null;

            }

            //Caso não conectar na planilha retorna um valor nulo e avisa.
            catch
            {
                TaskDialog.Show("Erro", "Não foi possivel connectar com a Planilha");
                return null;
            }

            //depois de executar o "try" com sucesso, ele finaliza a connexão com o excel.
            finally
            {
                connect.Close();
            }

        }
        #endregion

        #region Leitura dos Parametros de Projeto
        public static List<ParametrosDeProjeto> ObterParametrosProjeto()
        {
            MinhasVariaveis.connectionStringBase = @"provider = Microsoft.ACE.OLEDB.12.0;";
            MinhasVariaveis.dataSourceBase = "data source =";
            MinhasVariaveis.connectionPropertiesBase = ";Extended Properties = 'Excel 12.0 Xml;HDR=YES;IMAX=1;'";


            string connectionString = MinhasVariaveis.connectionStringBase +
                                      MinhasVariaveis.dataSourceBase +
                                      MinhasVariaveis.path +
                                      MinhasVariaveis.connectionPropertiesBase;


            OleDbConnection connect = new OleDbConnection(connectionString);
            string comandoSql = "Select * from [Cabeçalho$]";
            OleDbCommand comando = new OleDbCommand(comandoSql, connect);
            List<ParametrosDeProjeto> listaCabeçalho = new List<ParametrosDeProjeto>();

            try
            {
                connect.Open();
                OleDbDataReader rd = comando.ExecuteReader();

                //depois de connectado ele vai ler as linhas enquanto houver informação.
                while (rd.Read())
                {
                    listaCabeçalho.Add(new ParametrosDeProjeto()
                    {
                        //Campos para Parametros do Projeto
                        PROJETO = rd["PROJETO"].ToString(),
                        FASEPROJETO = rd["FASE PROJETO"].ToString(),
                        ABREVIAÇÃOCLIENTE = rd["ABREVIAÇÃO CLIENTE"].ToString(),
                        ABREVIAÇÃOFASE = rd["ABREVIAÇÃO FASE"].ToString(),
                        NUMEROPROJETO = rd["NUMERO PROJETO"].ToString(),
                        ENDEREÇO = rd["ENDEREÇO"].ToString()
                    });
                }
                if (listaCabeçalho.Count > 0)
                {
                    return listaCabeçalho;
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                TaskDialog.Show("Erro", "Não foi possivel connectar com a Planilha");
                return null;
            }
            finally
            {
                connect.Close();
            }
        }
        #endregion

        #region Metodo para ativar o campo de revisão do carimbo
        public void CheckRevisão(UIDocument uidoc, Document doc)
        {
            
            List<Entidades> listaRevisão = ObterCamposExcel(MinhasVariaveis.disciplina);

            if (listaRevisão != null)
            {
                try
                {
                    foreach (var vars in listaRevisão)
                    {
                        ViewSheet viewSheet = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewSheet))
                            .OfCategory(BuiltInCategory.OST_Sheets)
                            .Cast<ViewSheet>()
                            .FirstOrDefault(vs => vs.SheetNumber == vars.FOLHA);

                        List<ViewSheet> vSheetList = new List<ViewSheet>();
                        vSheetList.Add(viewSheet);


                        IList<ElementId> vSheetId = new List<ElementId>();

                        vSheetId.Add(viewSheet.Id);


                        foreach (ViewSheet vsId in vSheetList)
                        {
                            List<BuiltInCategory> cats = new List<BuiltInCategory>();
                            cats.Add(BuiltInCategory.OST_TitleBlocks);


                            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(cats);

                            IList<Element> tElements = new FilteredElementCollector(doc).OwnedByView(viewSheet.Id)
                                .WherePasses(filter)
                                .WhereElementIsNotElementType()
                                .ToElements();


                            foreach (Element elem in tElements)
                            {
                                List<Parameter> lParam = new List<Parameter>();
                                lParam.Add(elem.LookupParameter("R01"));
                                lParam.Add(elem.LookupParameter("R02"));
                                lParam.Add(elem.LookupParameter("R03"));
                                lParam.Add(elem.LookupParameter("R04"));
                                lParam.Add(elem.LookupParameter("R05"));
                                lParam.Add(elem.LookupParameter("R06"));
                                
                                if (viewSheet.SheetNumber == vars.FOLHA)
                                {
                                    foreach (Parameter par in lParam)
                                    {
                                        #region Revisão 00
                                        if (vars.REVISÃO == "00" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 01
                                        if (vars.REVISÃO == "01" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 02
                                        if (vars.REVISÃO == "02" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 03
                                        if (vars.REVISÃO == "03" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 04
                                        if (vars.REVISÃO == "04" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(0);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 05
                                        if (vars.REVISÃO == "05" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(0);
                                            }
                                        }
                                        #endregion
                                        #region Revisão 06
                                        if (vars.REVISÃO == "06" && viewSheet.SheetNumber == vars.FOLHA)
                                        {
                                            if (par.Definition.Name.Equals("R01"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R02"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R03"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R04"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R05"))
                                            {
                                                par.Set(1);
                                            }
                                            if (par.Definition.Name.Equals("R06"))
                                            {
                                                par.Set(1);
                                            }
                                        }
                                        #endregion
                                    }

                                }

                            }

                        }
                    }

                }

                catch 
                {
                    
                }

            }

        }


        #endregion

        #region Variaveis criadas para varios Metodos
        public static class MinhasVariaveis
        {
            public static string stats { get; set; }
            public static string disciplina { get; set; }
            public static string dataSourceBase { get; set; }
            public static string path { get; set; }
            public static string connectionStringBase { get; set; }
            public static string connectionPropertiesBase { get; set; }

        }
        #endregion

        #region Lista Parametros de Projeto
        /// <summary>
        /// lista de entidades para para preencher o parametro de projetos
        /// </summary>
        public class ParametrosDeProjeto
        {
            public string PROJETO { get; set; }
            public string FASEPROJETO { get; set; }
            public string DISCIPLINA { get; set; }
            public string ABREVIAÇÃOCLIENTE { get; set; }
            public string ABREVIAÇÃOFASE { get; set; }
            public string NUMEROPROJETO { get; set; }
            public string DATAINICIAL { get; set; }
            public string ENDEREÇO { get; set; }
        }
        #endregion

        #region Lista de Entidades
        /// <summary>
        /// Lista de entidades para salvar da planilha excel.
        /// </summary>
        public class Entidades
        {
            public string FOLHA { get; set; }
            public string NOMEFOLHA { get; set; }
            public string TITULO01 { get; set; }
            public string TITULO02 { get; set; }
            public string TITULO03 { get; set; }
            public string TITULO04 { get; set; }
            public string REVISÃO { get; set; }
            public string ESCALA { get; set; }
            public string DATAR00 { get; set; }
            public string DESCRIÇÃOR00 { get; set; }
            public string DESENHOR00 { get; set; }
            public string VERIFICADOR00 { get; set; }
            public string DATAR01 { get; set; }
            public string DESCRIÇÃOR01 { get; set; }
            public string DESENHOR01 { get; set; }
            public string VERIFICADOR01 { get; set; }
            public string DATAR02 { get; set; }
            public string DESCRIÇÃOR02 { get; set; }
            public string DESENHOR02 { get; set; }
            public string VERIFICADOR02 { get; set; }
            public string DATAR03 { get; set; }
            public string DESCRIÇÃOR03 { get; set; }
            public string DESENHOR03 { get; set; }
            public string VERIFICADOR03 { get; set; }
            public string DATAR04 { get; set; }
            public string DESCRIÇÃOR04 { get; set; }
            public string DESENHOR04 { get; set; }
            public string VERIFICADOR04 { get; set; }
            public string DATAR05 { get; set; }
            public string DESCRIÇÃOR05 { get; set; }
            public string DESENHOR05 { get; set; }
            public string VERIFICADOR05 { get; set; }
            public string DATAR06 { get; set; }
            public string DESCRIÇÃOR06 { get; set; }
            public string DESENHOR06 { get; set; }
            public string VERIFICADOR06 { get; set; }

        }
        #endregion

    }
}
