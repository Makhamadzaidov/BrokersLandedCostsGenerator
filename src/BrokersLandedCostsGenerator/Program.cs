using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BrokersLandedCostsGenerator
{
    class Program
    {
        private static SAPbobsCOM.Company oCom;
        private static SAPbobsCOM.Recordset oRs;
        private static string path = "log.txt";
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

                oCom = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                oRs = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "992" & pVal.BeforeAction & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item((object)pVal.FormUID);
                SAPbouiCOM.Item tem1 = form.Items.Item((object)"2");
                SAPbouiCOM.Item tem2 = form.Items.Add("generate", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                tem2.Left = tem1.Left + tem1.Width + 10;
                tem2.Top = tem1.Top;
                tem2.Width = 100;
                tem2.Height = tem1.Height;
                ((SAPbouiCOM.IButton)tem2.Specific).Caption = "Закупка услуг брокеров";
            }
            if (!(pVal.FormTypeEx == "992" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "generate" & !pVal.BeforeAction))
                return;
            SAPbouiCOM.DBDataSource dbDataSource = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item((object)pVal.FormUID).DataSources.DBDataSources.Item((object)"OIPF");
            if (dbDataSource.GetValue((object)"DocEntry", 0) == "")
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Документ не найден");
            }
            else
            {
                if (File.Exists(Program.path))
                {
                    if (((IEnumerable<string>)File.ReadAllLines(Program.path)).ToList<string>().Contains(dbDataSource.GetValue((object)"DocEntry", 0)))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Аддон уже запускался для этого документа.");
                        return;
                    }
                }
                else
                    File.Create(Program.path);
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Начался процесс генерации документов закупки услуг брокеров.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbobsCOM.LandedCostsService businessService = (SAPbobsCOM.LandedCostsService)Program.oCom.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService);
                SAPbobsCOM.LandedCostParams dataInterface = (SAPbobsCOM.LandedCostParams)businessService.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams);
                dataInterface.LandedCostNumber = int.Parse(dbDataSource.GetValue((object)"DocEntry", 0));
                SAPbobsCOM.LandedCost landedCost = businessService.GetLandedCost(dataInterface);
                Program.oRs.DoQuery("SELECT \"AlcCode\", \"AlcName\", \"LaCAllcAcc\" FROM OALC ");
                bool flag1 = false;
                for (int vtIndex = 0; vtIndex < landedCost.LandedCost_CostLines.Count; ++vtIndex)
                {
                    bool flag2 = false;
                    if (landedCost.LandedCost_CostLines.Item((object)vtIndex).Broker != "")
                    {
                        SAPbobsCOM.Documents businessObject = (SAPbobsCOM.Documents)Program.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        businessObject.CardCode = landedCost.LandedCost_CostLines.Item((object)vtIndex).Broker;
                        businessObject.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                        Program.oRs.MoveFirst();
                        while (!Program.oRs.EoF)
                        {
                            if (Program.oRs.Fields.Item((object)"AlcCode").Value.ToString() == landedCost.LandedCost_CostLines.Item((object)vtIndex).LandedCostCode)
                            {
                                flag2 = true;
                                break;
                            }
                            Program.oRs.MoveNext();
                        }
                        if (!flag2)
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Не удалось найти брокера " + landedCost.LandedCost_CostLines.Item((object)vtIndex).Broker + " в справочнике");
                            return;
                        }
                        try
                        {
                            businessObject.Lines.ItemDescription = Program.oRs.Fields.Item((object)"AlcName").Value.ToString();
                            businessObject.Lines.AccountCode = Program.oRs.Fields.Item((object)"LaCAllcAcc").Value.ToString();
                            businessObject.Lines.LineTotal = landedCost.LandedCost_CostLines.Item((object)vtIndex).amount;
                            // businessObject.Lines.VatGroup = "B0"; //B0
                            businessObject.DocumentReferences.ReferencedDocEntry = landedCost.DocEntry;
                            businessObject.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_LandedCosts;
                            int errCode = businessObject.Add();
                            if (errCode != 0)
                            {
                                flag1 = true;
                                string errMsg;
                                Program.oCom.GetLastError(out errCode, out errMsg);
                                Console.WriteLine(errMsg);
                                Application.SBO_Application.StatusBar.SetText(string.Format("Ошибка при генерации закупки для {0} [{1}]: {2}", (object)landedCost.LandedCost_CostLines.Item((object)vtIndex).Broker, (object)errCode, (object)errMsg), Type: SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                            }
                            else
                                Application.SBO_Application.StatusBar.SetText("Генерация закупок услуг для брокера " + landedCost.LandedCost_CostLines.Item((object)vtIndex).Broker + " завершена", Type: SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
                if (flag1)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Произошли ошибки при генерации закупок брокеров. Просмотрите логи чтобы их увидеть. Генерации закупок услуг брокеров завершена.");
                }
                else
                {
                    using (StreamWriter streamWriter = File.AppendText(Program.path))
                        streamWriter.WriteLine(dbDataSource.GetValue((object)"DocEntry", 0));
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Генерации закупок услуг брокеров завершена", Type: SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
