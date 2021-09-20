#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Creation;
#endregion

namespace RevitAddin
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Criando uma aba para o comando.
            String tabName = "JMBF";
            application.CreateRibbonTab(tabName);

            //adicionando o "Ribbon Panel" a aba criada
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Comandos JMBF");

            //selecionando o caminho assembly para o comando
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;


            PushButtonData btn_Carimbo = new PushButtonData("Carimbo",
                                                        "Carimbo",
                                                        thisAssemblyPath,
                                                        "RevitAddin.CreateSheet");

            //PushButtonData btn_PipeTag = new PushButtonData("PipeTag",
            //                                                "PipeTag",
            //                                                thisAssemblyPath,
            //                                                "RevitAddin.PipeTags");

            PushButton pushButton_Carbimo = ribbonPanel.AddItem(btn_Carimbo) as PushButton;
            //PushButton pushButton_PipeTag = ribbonPanel.AddItem(btn_PipeTag) as PushButton;


            //pushButton_PipeTag.ToolTip = "Insere Tag nas Tubulações hidraulicas";
            pushButton_Carbimo.ToolTip = "Preenchimento do Carimbo pela planilha do Excel";

            

            //BitmapImage btn_PipeTagImg = new BitmapImage(new Uri(@"X:\12 - Revit Icons\PipeTag.png"));
            //pushButton_PipeTag.LargeImage = btn_PipeTagImg;

            BitmapImage btn_CarimboImg = new BitmapImage(new Uri(@"X:\12 - Revit Icons\carimbo.png"));
            pushButton_Carbimo.LargeImage = btn_CarimboImg;


            return Result.Succeeded;


        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
