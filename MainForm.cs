/*
 * Created by SharpDevelop.
 * User: benoit le guern
 * Date: 05/08/2008
 * Time: 17:17
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace SuperUploader
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		private string files = null;
		private int nbFiles = 0;
		private const string CONFIG_FILENAME = "config.xml";
		
		[STAThread]
		public static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			LoadConfigFile();
			UpdateXls();
			UpdateItem();
			UpdateCustomer();
			UpdateSupplier();
			UpdateRouting();
			UpdateProdStruct();
			UpdateWorkCenter();
			UpdateProductionLine();
			UpdateMeasure();
		}
		
		private void LoadConfigFile()
		{				
			string xls_folder, 
			item_file, 
			item_site_cell_prod_line_file, 
			item_prod_line_file, 
			item_leader_file, 
			item_analysis_code_file,
			item_analysis_code_brand_file,
			item_intrastat_code_file,
			item_intrastat_file,
			item_raw_file, 
			item_dsrp_file,
			item_cost_file,
			item_v9_file, 
			item_v9_prod_line_file, 
			item_v9_last_prod_line_file, 
			item_v9_cost_file,
			item_general_params_file,
			item_comment_fc_file,
			item_comment_fb_file,
			item_comment_wa_file,
			item_logistics_file,
			item_v9_logistics_file,
			customer_business_relation_file, 
			customer_financial_file, 
			customer_file,
			customer_delivery_file,
			customer_tree_file,
			customer_item_file,
			customer_general_params_file,
			supplier_business_relation_file,
			supplier_financial_file,
			supplier_file,
			supplier_item_file,
			supplier_v9_file,
			supplier_v9_item_file,
			supplier_code_v9_qad2008_file,
			supplier_general_params_file,
			supplier_pricing_file,
			routing_file,
			routing_v9_file,
			routing_comment_file,
			routing_comment_v9_file,
			prod_struct_file,
			prod_struct_v9_file,
			prod_struct_code_file,
			prod_struct_code_v9_file,
			work_center_file,
			work_center_v9_file,
			production_line_file,
			measure_file,
			measure_v9_file;
			
			if (File.Exists(System.Windows.Forms.Application.StartupPath + "\\" + CONFIG_FILENAME))
			{
				try
	            {
	                XmlDocument config = new XmlDocument();
	
	                config.Load(System.Windows.Forms.Application.StartupPath + "\\" + CONFIG_FILENAME);
	
	                xls_folder = config.DocumentElement.SelectSingleNode("//config/directories/directory[@name='xls_folder']/@value").Value;
			
	                item_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_file']/@value").Value;
	                item_site_cell_prod_line_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_site_cell_prod_line_file']/@value").Value;
	                item_prod_line_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_prod_line_file']/@value").Value;
	                item_leader_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_leader_file']/@value").Value;
	                item_analysis_code_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_analysis_code_file']/@value").Value;
	                item_analysis_code_brand_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_analysis_code_brand_file']/@value").Value;
	                item_intrastat_code_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_intrastat_code_file']/@value").Value;
	                item_intrastat_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_intrastat_file']/@value").Value;
	                item_raw_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_raw_file']/@value").Value;
	                item_dsrp_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_dsrp_file']/@value").Value;
	                item_cost_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_cost_file']/@value").Value;
	      			item_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_v9_file']/@value").Value;
	                item_v9_prod_line_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_v9_prod_line_file']/@value").Value;
	                item_v9_last_prod_line_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_v9_last_prod_line_file']/@value").Value;
	                item_v9_cost_file  = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_v9_cost_file']/@value").Value;
					item_general_params_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_general_params_file']/@value").Value;
					item_comment_fc_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_comment_fc_file']/@value").Value;
					item_comment_fb_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_comment_fb_file']/@value").Value;
					item_comment_wa_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_comment_wa_file']/@value").Value;
			
					item_logistics_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_logistics_file']/@value").Value;
					item_v9_logistics_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='item_v9_logistics_file']/@value").Value;
			
					
	                customer_business_relation_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_business_relation_file']/@value").Value;
					customer_financial_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_financial_file']/@value").Value;
					customer_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_file']/@value").Value;
					customer_delivery_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_delivery_file']/@value").Value;
					customer_tree_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_tree_file']/@value").Value;
	                customer_item_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_item_file']/@value").Value;
	                customer_general_params_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='customer_general_params_file']/@value").Value;
				
	                supplier_business_relation_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_business_relation_file']/@value").Value;
					supplier_financial_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_financial_file']/@value").Value;
					supplier_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_file']/@value").Value;
					supplier_item_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_item_file']/@value").Value;
					supplier_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_v9_file']/@value").Value;
					supplier_v9_item_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_v9_item_file']/@value").Value;
					supplier_code_v9_qad2008_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_code_v9_qad2008_file']/@value").Value;
					supplier_general_params_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_general_params_file']/@value").Value;
	            	supplier_pricing_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='supplier_pricing_file']/@value").Value;
	            		
					routing_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='routing_file']/@value").Value;
	                routing_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='routing_v9_file']/@value").Value;
	                routing_comment_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='routing_comment_file']/@value").Value;
	                routing_comment_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='routing_comment_v9_file']/@value").Value;
	                
	                prod_struct_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='prod_struct_file']/@value").Value;
	                prod_struct_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='prod_struct_v9_file']/@value").Value;
	                prod_struct_code_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='prod_struct_code_file']/@value").Value;
	                prod_struct_code_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='prod_struct_code_v9_file']/@value").Value;
	                
	                work_center_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='work_center_file']/@value").Value;
	                work_center_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='work_center_v9_file']/@value").Value;
	               
	                production_line_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='production_line_file']/@value").Value;
	               
	                measure_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='measure_file']/@value").Value;
	               	measure_v9_file = config.DocumentElement.SelectSingleNode("//config/files/file[@name='measure_v9_file']/@value").Value;
	               
	                if ((xls_folder != null) && (Directory.Exists(xls_folder)))
	            	{
						textBoxXlsFolder.Text = xls_folder;
					}
					
					if ((item_file != null) && (File.Exists(item_file)))
		            {
		            	textBoxItemFile.Text = item_file;
		            }
					
					if ((item_site_cell_prod_line_file != null) && (File.Exists(item_site_cell_prod_line_file)))
		            {
		            	textBoxItemSiteCellProdLineFile.Text = item_site_cell_prod_line_file;
		            }
					
					if ((item_prod_line_file != null) && (File.Exists(item_prod_line_file)))
		            {
		            	textBoxItemProdLineFile.Text = item_prod_line_file;
		            }
					
					if ((item_leader_file != null) && (File.Exists(item_leader_file)))
		            {
		            	textBoxItemLeaderFile.Text = item_leader_file;
		            }
					
					if ((item_analysis_code_file != null) && (File.Exists(item_analysis_code_file)))
		            {
		            	textBoxItemAnalysisCodeFile.Text = item_analysis_code_file;
		            }
					
					if ((item_analysis_code_brand_file != null) && (File.Exists(item_analysis_code_brand_file)))
		            {
		            	textBoxItemAnalysisCodeBrandFile.Text = item_analysis_code_brand_file;
		            }
					
					if ((item_intrastat_code_file != null) && (File.Exists(item_intrastat_code_file)))
		            {
		            	textBoxItemIntrastatCodeFile.Text = item_intrastat_code_file;
		            }
					
					if ((item_intrastat_file != null) && (File.Exists(item_intrastat_file)))
		            {
		            	textBoxItemIntrastatFile.Text = item_intrastat_file;
		            }
					
					if ((item_raw_file != null) && (File.Exists(item_raw_file)))
		            {
		            	textBoxItemRawFile.Text = item_raw_file;
		            }
					
					if ((item_dsrp_file != null) && (File.Exists(item_dsrp_file)))
		            {
		            	textBoxItemDSRPFile.Text = item_dsrp_file;
		            }
					
					if ((item_cost_file != null) && (File.Exists(item_cost_file)))
		            {
		            	textBoxItemCostFile.Text = item_cost_file;
		            }
					
					if ((item_v9_file != null) && (File.Exists(item_v9_file)))
		            {
		            	textBoxItemV9File.Text = item_v9_file;
		            }
					
					if ((item_v9_prod_line_file != null) && (File.Exists(item_v9_prod_line_file)))
		            {
		            	textBoxItemV9ProdLineFile.Text = item_v9_prod_line_file;
		            }
					
					if ((item_v9_last_prod_line_file != null) && (File.Exists(item_v9_last_prod_line_file)))
		            {
		            	textBoxItemV9LastProdLineFile.Text = item_v9_last_prod_line_file;
		            }
					
					if ((item_v9_cost_file != null) && (File.Exists(item_v9_cost_file)))
		            {
		            	textBoxItemV9CostFile.Text = item_v9_cost_file;
		            }
					
					if ((item_general_params_file != null) && (File.Exists(item_general_params_file)))
		            {
		            	textBoxItemGeneralParamsFile.Text = item_general_params_file;
		            }
					
					
					if ((item_comment_fc_file != null) && (File.Exists(item_comment_fc_file)))
		            {
		            	textBoxItemCommentFCFile.Text = item_comment_fc_file;
		            }
					
					if ((item_comment_fb_file != null) && (File.Exists(item_comment_fb_file)))
		            {
		            	textBoxItemCommentFBFile.Text = item_comment_fb_file;
		            }
					
					if ((item_comment_wa_file != null) && (File.Exists(item_comment_wa_file)))
		            {
		            	textBoxItemCommentWAFile.Text = item_comment_wa_file;
		            }
					
					if ((item_logistics_file != null) && (File.Exists(item_logistics_file)))
		            {
		            	textBoxItemLogisticsFile.Text = item_logistics_file;
		            }
					
					if ((item_v9_logistics_file != null) && (File.Exists(item_v9_logistics_file)))
		            {
		            	textBoxItemV9LogisticsFile.Text = item_v9_logistics_file;
		            }
					
					
					if ((customer_business_relation_file != null) && (File.Exists(customer_business_relation_file)))
		            {
		            	textBoxCustomerBusinessRelationFile.Text = customer_business_relation_file;
		            }
					
					if ((customer_financial_file != null) && (File.Exists(customer_financial_file)))
		            {
		            	textBoxCustomerFinancialFile.Text = customer_financial_file;
		            }
					
					if ((customer_file != null) && (File.Exists(customer_file)))
		            {
		            	textBoxCustomerFile.Text = customer_file;
		            }
					
					if ((customer_delivery_file != null) && (File.Exists(customer_delivery_file)))
		            {
		            	textBoxCustomerDeliveryFile.Text = customer_delivery_file;
		            }
					
					if ((customer_tree_file != null) && (File.Exists(customer_tree_file)))
		            {
		            	textBoxCustomerTreeFile.Text = customer_tree_file;
		            }
					
					if ((customer_item_file != null) && (File.Exists(customer_item_file)))
		            {
		            	textBoxCustomerItemFile.Text = customer_item_file;
		            }
					
					if ((customer_general_params_file != null) && (File.Exists(customer_general_params_file)))
		            {
		            	textBoxCustomerGeneralParamsFile.Text = customer_general_params_file;
		            }
					
					
					if ((supplier_business_relation_file != null) && (File.Exists(supplier_business_relation_file)))
		            {
		            	textBoxSupplierBusinessRelationFile.Text = supplier_business_relation_file;
		            }
					
					if ((supplier_financial_file != null) && (File.Exists(supplier_financial_file)))
		            {
		            	textBoxSupplierFinancialFile.Text = supplier_financial_file;
		            }
					
					if ((supplier_file != null) && (File.Exists(supplier_file)))
		            {
		            	textBoxSupplierFile.Text = supplier_file;
		            }
					
					if ((supplier_item_file != null) && (File.Exists(supplier_item_file)))
		            {
		            	textBoxSupplierItemFile.Text = supplier_item_file;
		            }
					
					if ((supplier_v9_file != null) && (File.Exists(supplier_v9_file)))
		            {
		            	textBoxSupplierV9File.Text = supplier_v9_file;
		            }
					
					if ((supplier_v9_item_file != null) && (File.Exists(supplier_v9_item_file)))
		            {
		            	textBoxSupplierV9ItemFile.Text = supplier_v9_item_file;
		            }
					
					if ((supplier_code_v9_qad2008_file != null) && (File.Exists(supplier_code_v9_qad2008_file)))
		            {
		            	textBoxSupplierCodeV9QAD2008File.Text = supplier_code_v9_qad2008_file;
		            }
					
					if ((supplier_general_params_file != null) && (File.Exists(supplier_general_params_file)))
		            {
		            	textBoxSupplierGeneralParamsFile.Text = supplier_general_params_file;
		            }
					
					if ((supplier_pricing_file != null) && (File.Exists(supplier_pricing_file)))
		            {
		            	textBoxSupplierPricingFile.Text = supplier_pricing_file;
		            }
					
					
					if ((routing_file != null) && (File.Exists(routing_file)))
		            {
		            	textBoxRoutingFile.Text = routing_file;
		            }
					
					if ((routing_v9_file != null) && (File.Exists(routing_v9_file)))
		            {
		            	textBoxRoutingV9File.Text = routing_v9_file;
		            }
					
					if ((routing_comment_file != null) && (File.Exists(routing_comment_file)))
		            {
		            	textBoxRoutingCommentFile.Text = routing_comment_file;
		            }
					
					if ((routing_comment_v9_file != null) && (File.Exists(routing_comment_v9_file)))
		            {
		            	textBoxRoutingCommentV9File.Text = routing_comment_v9_file;
		            }
					
					
					if ((prod_struct_file != null) && (File.Exists(prod_struct_file)))
		            {
		            	textBoxProdStructFile.Text = prod_struct_file;
		            }
					if ((prod_struct_v9_file != null) && (File.Exists(prod_struct_v9_file)))
		            {
		            	textBoxProdStructV9File.Text = prod_struct_v9_file;
		            }
					if ((prod_struct_code_file != null) && (File.Exists(prod_struct_code_file)))
		            {
		            	textBoxProdStructCodeFile.Text = prod_struct_code_file;
		            }
					if ((prod_struct_code_v9_file != null) && (File.Exists(prod_struct_code_v9_file)))
		            {
		            	textBoxProdStructCodeV9File.Text = prod_struct_code_v9_file;
		            }
					
					
					if ((work_center_file != null) && (File.Exists(work_center_file)))
		            {
		            	textBoxWorkCenterFile.Text = work_center_file;
		            }
					
					if ((work_center_v9_file != null) && (File.Exists(work_center_v9_file)))
		            {
		            	textBoxWorkCenterV9File.Text = work_center_v9_file;
		            }
					
					
					if ((production_line_file != null) && (File.Exists(production_line_file)))
		            {
		            	textBoxProductionLineFile.Text = production_line_file;
		            }
					
					
					if ((measure_file != null) && (File.Exists(measure_file)))
		            {
		            	textBoxMeasureFile.Text = measure_file;
		            }
					
					if ((measure_v9_file != null) && (File.Exists(measure_v9_file)))
		            {
		            	textBoxMeasureV9File.Text = measure_v9_file;
		            }
				}
	            catch (Exception)
	            {
	            	
	            }
			}
		}
		
		private void UpdateXls()
		{
			/* INIT */
			files = null;
			nbFiles = 0;
			buttonXls.Enabled = false;
			
			if (Directory.Exists(textBoxXlsFolder.Text))
			{
				string [] tmpFiles = Directory.GetFiles(textBoxXlsFolder.Text);
				
				foreach (string tmpFile in tmpFiles)
				{
					if (Path.GetExtension(tmpFile).ToLower().Equals(".xls"))
					{
						files += tmpFile.ToLower() + ";";
						nbFiles++;
					}
				}
			}
			
			if (nbFiles > 0)
			{
				buttonXls.Enabled = true;
			}
		}
		
		private void UpdateItem()
		{
			if ((textBoxItemFile.Text != "") &&
			    (textBoxItemSiteCellProdLineFile.Text != "") &&
			    (textBoxItemProdLineFile.Text != "") && 
			    (textBoxItemAnalysisCodeFile.Text != "") &&
			    (textBoxItemAnalysisCodeBrandFile.Text != "") &&
			    (textBoxItemIntrastatFile.Text != "") &&
			    (textBoxItemIntrastatCodeFile.Text != "") &&
			    (textBoxItemLeaderFile.Text != "") &&
			    (textBoxItemRawFile.Text != "") &&
			    (textBoxItemV9File.Text != "") && 
			    (textBoxItemV9ProdLineFile.Text != "") && 
			    (textBoxItemV9LastProdLineFile.Text != "") &&
			    (textBoxItemCommentFCFile.Text != "") && 
			    (textBoxItemCommentFBFile.Text != "") && 
			    (textBoxItemCommentWAFile.Text != "") && 
			    (textBoxItemLogisticsFile.Text != "") && 
			    (textBoxItemV9LogisticsFile.Text != ""))
			{
				buttonItem.Enabled = true;
			}
			else
			{
				buttonItem.Enabled = false;
			}
		}
		
		private void UpdateCustomer()
		{
			if ((textBoxCustomerBusinessRelationFile.Text != "") &&
			    (textBoxCustomerFile.Text != "") &&
			    (textBoxCustomerFinancialFile.Text != "") && 
			    (textBoxCustomerDeliveryFile.Text != "") &&
			    (textBoxCustomerTreeFile.Text != "") &&
			    (textBoxCustomerItemFile.Text != ""))
			{
				buttonCustomer.Enabled = true;
			}
			else
			{
				buttonCustomer.Enabled = false;
			}
		}
		
		private void UpdateSupplier()
		{
			if ((textBoxSupplierBusinessRelationFile.Text != "") &&
			    (textBoxSupplierFile.Text != "") &&
			    (textBoxSupplierItemFile.Text != "") && 
			    (textBoxSupplierFinancialFile.Text != "") && 
			    (textBoxSupplierV9File.Text != "") && 
			    (textBoxSupplierV9ItemFile.Text != "") && 
			    (textBoxSupplierGeneralParamsFile.Text != "") && 
			    (textBoxSupplierCodeV9QAD2008File.Text != "") && 
			    (textBoxSupplierPricingFile.Text != ""))
			{
				buttonSupplier.Enabled = true;
				buttonSupplierPricing.Enabled = true;
			}
			else
			{
				buttonSupplier.Enabled = false;
				buttonSupplierPricing.Enabled = false;
			}
		}
		
		private void UpdateRouting()
		{
			if ((textBoxRoutingFile.Text != "") &&
			    (textBoxRoutingV9File.Text != "") &&
			    (textBoxRoutingCommentFile.Text != "") &&
			    (textBoxRoutingCommentV9File.Text != ""))
			{
				buttonRouting.Enabled = true;
			}
			else
			{
				buttonRouting.Enabled = false;
			}
		}
		
		private void UpdateProdStruct()
		{
			if ((textBoxProdStructFile.Text != "") &&
			    (textBoxProdStructV9File.Text != "")&&
			    (textBoxProdStructCodeFile.Text != "")&&
			    (textBoxProdStructCodeV9File.Text != ""))
			{
				buttonProdStruct.Enabled = true;
			}
			else
			{
				buttonProdStruct.Enabled = false;
			}
		}
		
		void UpdateWorkCenter()
		{
			if ((textBoxWorkCenterFile.Text != "") &&
			    (textBoxWorkCenterV9File.Text != ""))
			{
				buttonWorkCenter.Enabled = true;
			}
			else
			{
				buttonWorkCenter.Enabled = false;
			}
		}
		
		void UpdateProductionLine()
		{
			if (textBoxProductionLineFile.Text != "")
			{
				buttonProductionLine.Enabled = true;
			}
			else
			{
				buttonProductionLine.Enabled = false;
			}
		}
		
		void ButtonQuitClick(object sender, EventArgs e)
		{
			try
            {
                XmlDocument config = new XmlDocument();			
			
                config.LoadXml("<?xml version=\"1.0\"?>" +
                               "<config>" +
                                "<directories>" +
                                    "<directory name=\"xls_folder\" value=\"" + textBoxXlsFolder.Text + "\" />" +
                               "</directories>" + 
                                "<files>" +
                                	"<file name=\"item_file\" value=\"" + textBoxItemFile.Text + "\" />" +
                                	"<file name=\"item_site_cell_prod_line_file\" value=\"" + textBoxItemSiteCellProdLineFile.Text + "\" />" +
                                	"<file name=\"item_prod_line_file\" value=\"" + textBoxItemProdLineFile.Text + "\" />" +
                                	"<file name=\"item_leader_file\" value=\"" + textBoxItemLeaderFile.Text + "\" />" +
                                	"<file name=\"item_analysis_code_file\" value=\"" + textBoxItemAnalysisCodeFile.Text + "\" />" +
                                	"<file name=\"item_analysis_code_brand_file\" value=\"" + textBoxItemAnalysisCodeBrandFile.Text + "\" />" +
                                	"<file name=\"item_intrastat_code_file\" value=\"" + textBoxItemIntrastatCodeFile.Text + "\" />" +
                                	"<file name=\"item_intrastat_file\" value=\"" + textBoxItemIntrastatFile.Text + "\" />" +
                                	"<file name=\"item_raw_file\" value=\"" + textBoxItemRawFile.Text + "\" />" +
                                	"<file name=\"item_dsrp_file\" value=\"" + textBoxItemDSRPFile.Text + "\" />" +
                                	"<file name=\"item_cost_file\" value=\"" + textBoxItemCostFile.Text + "\" />" +
                                	"<file name=\"item_v9_file\" value=\"" + textBoxItemV9File.Text + "\" />" +
                                	"<file name=\"item_v9_prod_line_file\" value=\"" + textBoxItemV9ProdLineFile.Text + "\" />" +
                                	"<file name=\"item_v9_last_prod_line_file\" value=\"" + textBoxItemV9LastProdLineFile.Text + "\" />" +
                                	"<file name=\"item_v9_cost_file\" value=\"" + textBoxItemV9CostFile.Text + "\" />" +
                                	"<file name=\"item_general_params_file\" value=\"" + textBoxItemGeneralParamsFile.Text + "\" />" +
                                	"<file name=\"item_comment_fc_file\" value=\"" + textBoxItemCommentFCFile.Text + "\" />" +
                                	"<file name=\"item_comment_fb_file\" value=\"" + textBoxItemCommentFBFile.Text + "\" />" +
                                	"<file name=\"item_comment_wa_file\" value=\"" + textBoxItemCommentWAFile.Text + "\" />" +
                                	"<file name=\"item_logistics_file\" value=\"" + textBoxItemLogisticsFile.Text + "\" />" +
                                	"<file name=\"item_v9_logistics_file\" value=\"" + textBoxItemV9LogisticsFile.Text + "\" />" +
                                	"<file name=\"customer_business_relation_file\" value=\"" + textBoxCustomerBusinessRelationFile.Text + "\" />" +
                                	"<file name=\"customer_financial_file\" value=\"" + textBoxCustomerFinancialFile.Text + "\" />" +
                                	"<file name=\"customer_file\" value=\"" + textBoxCustomerFile.Text + "\" />" +
                                	"<file name=\"customer_delivery_file\" value=\"" + textBoxCustomerDeliveryFile.Text + "\" />" +
                                	"<file name=\"customer_tree_file\" value=\"" + textBoxCustomerTreeFile.Text + "\" />" +
                                	"<file name=\"customer_item_file\" value=\"" + textBoxCustomerItemFile.Text + "\" />" +
                                	"<file name=\"customer_general_params_file\" value=\"" + textBoxCustomerGeneralParamsFile.Text + "\" />" +
                                	"<file name=\"supplier_business_relation_file\" value=\"" + textBoxSupplierBusinessRelationFile.Text + "\" />" +
                                	"<file name=\"supplier_financial_file\" value=\"" + textBoxSupplierFinancialFile.Text + "\" />" +
                                	"<file name=\"supplier_file\" value=\"" + textBoxSupplierFile.Text + "\" />" +
                                	"<file name=\"supplier_item_file\" value=\"" + textBoxSupplierItemFile.Text + "\" />" +
                                	"<file name=\"supplier_v9_file\" value=\"" + textBoxSupplierV9File.Text + "\" />" +
                                	"<file name=\"supplier_v9_item_file\" value=\"" + textBoxSupplierV9ItemFile.Text + "\" />" +
                                	"<file name=\"supplier_code_v9_qad2008_file\" value=\"" + textBoxSupplierCodeV9QAD2008File.Text + "\" />" +
                                	"<file name=\"supplier_general_params_file\" value=\"" + textBoxSupplierGeneralParamsFile.Text + "\" />" +
                                	"<file name=\"supplier_pricing_file\" value=\"" + textBoxSupplierPricingFile.Text + "\" />" +
                                	"<file name=\"routing_file\" value=\"" + textBoxRoutingFile.Text + "\" />" +
                                	"<file name=\"routing_v9_file\" value=\"" + textBoxRoutingV9File.Text + "\" />" +
                                	"<file name=\"routing_comment_file\" value=\"" + textBoxRoutingCommentFile.Text + "\" />" +
                                	"<file name=\"routing_comment_v9_file\" value=\"" + textBoxRoutingCommentV9File.Text + "\" />" +
                                	"<file name=\"prod_struct_file\" value=\"" + textBoxProdStructFile.Text + "\" />" +
                                	"<file name=\"prod_struct_v9_file\" value=\"" + textBoxProdStructV9File.Text + "\" />" +
                                	"<file name=\"prod_struct_code_file\" value=\"" + textBoxProdStructCodeFile.Text + "\" />" +
                                	"<file name=\"prod_struct_code_v9_file\" value=\"" + textBoxProdStructCodeV9File.Text + "\" />" +
                                	"<file name=\"work_center_file\" value=\"" + textBoxWorkCenterFile.Text + "\" />" +
                                	"<file name=\"work_center_v9_file\" value=\"" + textBoxWorkCenterV9File.Text + "\" />" +
                                	"<file name=\"production_line_file\" value=\"" + textBoxProductionLineFile.Text + "\" />" +
                                	"<file name=\"measure_file\" value=\"" + textBoxMeasureFile.Text + "\" />" +
                                	"<file name=\"measure_v9_file\" value=\"" + textBoxMeasureV9File.Text + "\" />" +
                                "</files>" + 
                               "</config>");

                config.Save(System.Windows.Forms.Application.StartupPath + "\\" + CONFIG_FILENAME);
            }
            catch (Exception)
            {
            	
            }
            
            this.Close();
            System.Windows.Forms.Application.Exit();
		}
		
		void ButtonOpenXlsFolderClick(object sender, EventArgs e)
		{
			folderBrowserDialogXls = new FolderBrowserDialog();
			
			if (textBoxXlsFolder.Text != "")
			{
				folderBrowserDialogXls.SelectedPath = textBoxXlsFolder.Text;
			}
			
			if (folderBrowserDialogXls.ShowDialog() == DialogResult.OK)
			{
				textBoxXlsFolder.Text = folderBrowserDialogXls.SelectedPath;
				
				UpdateXls();
			}
		}
		
		void ButtonXlsClick(object sender, EventArgs e)
		{
			buttonOpenXlsFolder.Enabled = false;
			buttonXls.Enabled = false;
			
			if (nbFiles > 0)
			{
				string [] filenames = files.Split(";".ToCharArray());
				
				foreach (string filename in filenames)
				{
					if (filename != "")
					{
						Build.ConvertXlsToCsv(filename);
					}
				}
			}
			
			buttonXls.Enabled = true;
			buttonOpenXlsFolder.Enabled = true;
		}
		
		void ButtonOpenItemFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files|*.csv";
			
			if (textBoxItemFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemSiteCellProdLineFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files|*.csv";
			
			if (textBoxItemSiteCellProdLineFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemSiteCellProdLineFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemSiteCellProdLineFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemProdLineFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemProdLineFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemProdLineFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemProdLineFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemLeaderFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemLeaderFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemLeaderFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemLeaderFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemRawFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemRawFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemRawFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemRawFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemV9ProdLineFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemV9ProdLineFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemV9ProdLineFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemV9ProdLineFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemV9LastProdLineFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemV9LastProdLineFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemV9LastProdLineFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemV9LastProdLineFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemAnalysisCodeFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemAnalysisCodeFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemAnalysisCodeFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemAnalysisCodeFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemIntrastatFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemIntrastatFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemIntrastatFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemIntrastatFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemDSRPFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemDSRPFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemDSRPFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemDSRPFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemGeneralParamsFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemGeneralParamsFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemGeneralParamsFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemGeneralParamsFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonItemClick(object sender, EventArgs e)
		{
			buttonOpenItemFile.Enabled = false;
			buttonOpenItemSiteCellProdLineFile.Enabled = false;
			buttonOpenItemProdLineFile.Enabled = false;
			buttonOpenItemLeaderFile.Enabled = false;
			buttonOpenItemAnalysisCodeFile.Enabled = false;
			buttonOpenItemAnalysisCodeBrandFile.Enabled = false;
			buttonOpenItemIntrastatCodeFile.Enabled = false;
			buttonOpenItemIntrastatFile.Enabled = false;
			buttonOpenItemRawFile.Enabled = false;
			buttonOpenItemCostFile.Enabled = false;
			buttonOpenItemV9File.Enabled = false;
			buttonOpenItemV9ProdLineFile.Enabled = false;
			buttonOpenItemV9LastProdLineFile.Enabled = false;
			buttonOpenItemV9CostFile.Enabled = false;
			buttonOpenItemDSRPFile.Enabled = false;
			buttonOpenItemGeneralParamsFile.Enabled = false;
			buttonOpenItemCommentFCFile.Enabled = false;
			buttonOpenItemCommentFBFile.Enabled = false;
			buttonOpenItemCommentWAFile.Enabled = false;
			buttonOpenItemLogisticsFile.Enabled = false;
			buttonOpenItemV9LogisticsFile.Enabled = false;
			buttonItem.Enabled = false;
			
			System.Data.DataTable itemFile = Build.GetDataTableFromCsvFile(textBoxItemFile.Text);
			System.Data.DataTable itemSiteCellProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemSiteCellProdLineFile.Text);
			System.Data.DataTable itemProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemProdLineFile.Text);
			System.Data.DataTable itemLeaderFile = Build.GetDataTableFromCsvFile(textBoxItemLeaderFile.Text);
			System.Data.DataTable itemAnalysisCodeFile = Build.GetDataTableFromCsvFile(textBoxItemAnalysisCodeFile.Text);
			System.Data.DataTable itemAnalysisCodeBrandFile = Build.GetDataTableFromCsvFile(textBoxItemAnalysisCodeBrandFile.Text);
			System.Data.DataTable itemIntrastatCodeFile = Build.GetDataTableFromCsvFile(textBoxItemIntrastatCodeFile.Text);
			System.Data.DataTable itemIntrastatFile = Build.GetDataTableFromCsvFile(textBoxItemIntrastatFile.Text);
			System.Data.DataTable itemRawFile = Build.GetDataTableFromCsvFile(textBoxItemRawFile.Text);
			System.Data.DataTable itemCostFile = Build.GetDataTableFromCsvFile(textBoxItemCostFile.Text);
			System.Data.DataTable itemV9File = Build.GetDataTableFromCsvFile(textBoxItemV9File.Text);
			System.Data.DataTable itemV9ProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9ProdLineFile.Text);
			System.Data.DataTable itemV9LastProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9LastProdLineFile.Text);
			System.Data.DataTable itemV9CostFile = Build.GetDataTableFromCsvFile(textBoxItemV9CostFile.Text);
			System.Data.DataTable itemDSRPFile = Build.GetDataTableFromCsvFile(textBoxItemDSRPFile.Text);
			System.Data.DataTable itemGeneralParamsFile = Build.GetDataTableFromCsvFile(textBoxItemGeneralParamsFile.Text);
			System.Data.DataTable itemCommentFCFile = Build.GetDataTableFromCsvFile(textBoxItemCommentFCFile.Text);
			System.Data.DataTable itemCommentFBFile = Build.GetDataTableFromCsvFile(textBoxItemCommentFBFile.Text);
			System.Data.DataTable itemCommentWAFile = Build.GetDataTableFromCsvFile(textBoxItemCommentWAFile.Text);
			System.Data.DataTable itemLogisticsFile = Build.GetDataTableFromCsvFile(textBoxItemLogisticsFile.Text);
			System.Data.DataTable itemV9LogisticsFile = Build.GetDataTableFromCsvFile(textBoxItemV9LogisticsFile.Text);
			
			System.Data.DataTable itemTable = Build.Write141_Items(itemFile, 
					                                               itemSiteCellProdLineFile,
					                                               itemProdLineFile,
												                   itemLeaderFile, 
												                   itemRawFile,
												                   itemV9File,
												                   itemV9ProdLineFile, 
												                   itemV9LastProdLineFile,
												                   itemDSRPFile);
			if (itemTable != null)
			{
				Build.Write36213_Items(itemGeneralParamsFile, itemTable);
			}
			
			if (itemTable != null)
			{
				Build.Write1415_Items(itemRawFile, itemV9CostFile, itemTable);
			}
			
			if (itemTable != null)
			{
				Build.Write13_Items(itemLogisticsFile, itemV9LogisticsFile, itemTable);
			}
			
			Build.Write29223_Items(itemIntrastatCodeFile);
			
			if (itemTable != null)
			{
				Build.Write29226_Items(itemIntrastatFile, itemTable);
			}
			
			Build.WriteAnalysisCode_Items(itemAnalysisCodeFile, itemAnalysisCodeBrandFile, itemTable);
			
			Build.Write112_Items(itemCommentFCFile, itemCommentFBFile, itemCommentWAFile);
			
			buttonOpenItemFile.Enabled = true;
			buttonOpenItemSiteCellProdLineFile.Enabled = true;
			buttonOpenItemProdLineFile.Enabled = true;
			buttonOpenItemLeaderFile.Enabled = true;
			buttonOpenItemAnalysisCodeFile.Enabled = true;
			buttonOpenItemAnalysisCodeBrandFile.Enabled = true;
			buttonOpenItemIntrastatCodeFile.Enabled = true;
			buttonOpenItemIntrastatFile.Enabled = true;
			buttonOpenItemRawFile.Enabled = true;
			buttonOpenItemCostFile.Enabled = true;
			buttonOpenItemV9File.Enabled = true;
			buttonOpenItemV9ProdLineFile.Enabled = true;
			buttonOpenItemV9LastProdLineFile.Enabled = true;
			buttonOpenItemV9CostFile.Enabled = true;
			buttonOpenItemDSRPFile.Enabled = true;
			buttonOpenItemGeneralParamsFile.Enabled = true;
			buttonOpenItemCommentFCFile.Enabled = true;
			buttonOpenItemCommentFBFile.Enabled = true;
			buttonOpenItemCommentWAFile.Enabled = true;
			buttonOpenItemLogisticsFile.Enabled = true;
			buttonOpenItemV9LogisticsFile.Enabled = true;
			buttonItem.Enabled = true;
		}
		
		void ButtonOpenCustomerBusinessRelationFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerBusinessRelationFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerBusinessRelationFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerBusinessRelationFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenCustomerFinancialFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerFinancialFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerFinancialFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerFinancialFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenCustomerFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenCustomerDeliveryFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerDeliveryFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerDeliveryFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerDeliveryFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenCustomerTreeFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerTreeFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerTreeFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerTreeFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenCustomerItemFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerItemFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerItemFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerItemFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonCustomerClick(object sender, EventArgs e)
		{
			buttonOpenCustomerBusinessRelationFile.Enabled = false;
			buttonOpenCustomerFinancialFile.Enabled = false;
			buttonOpenCustomerFile.Enabled = false;
			buttonOpenCustomerDeliveryFile.Enabled = false;
			buttonOpenCustomerTreeFile.Enabled = false;
			buttonOpenCustomerItemFile.Enabled = false;
			buttonOpenCustomerGeneralParamsFile.Enabled = false;
			buttonCustomer.Enabled = false;
			
			System.Data.DataTable customerBusinessRelationFile = Build.GetDataTableFromCsvFile(textBoxCustomerBusinessRelationFile.Text);
			System.Data.DataTable customerFinancialFile = Build.GetDataTableFromCsvFile(textBoxCustomerFinancialFile.Text);
			System.Data.DataTable customerFile = Build.GetDataTableFromCsvFile(textBoxCustomerFile.Text);
			System.Data.DataTable customerDeliveryFile = Build.GetDataTableFromCsvFile(textBoxCustomerDeliveryFile.Text);
			System.Data.DataTable customerTreeFile = Build.GetDataTableFromCsvFile(textBoxCustomerTreeFile.Text);
			System.Data.DataTable customerItemFile = Build.GetDataTableFromCsvFile(textBoxCustomerItemFile.Text);
			System.Data.DataTable customerGeneralParamsFile = Build.GetDataTableFromCsvFile(textBoxCustomerGeneralParamsFile.Text);
			
			System.Data.DataTable businessRelationTable = Build.Write361431_Customers(customerBusinessRelationFile);
			System.Data.DataTable countyTable = Build.Write361331_Customers(customerBusinessRelationFile);
			System.Data.DataTable customerFinancialTable = Build.Write272011_Customers(customerFinancialFile);
			System.Data.DataTable customerTable = Build.Write211_Customers(customerFile);
			System.Data.DataTable customerDeliveryTable = Build.Write272021_Customers(customerDeliveryFile);
			System.Data.DataTable customerTreeTable = Build.WriteAnalysisCode_Customers(customerTreeFile);
			System.Data.DataTable customerItemTable = Build.Write115_Customers(customerItemFile);
			System.Data.DataTable customerGeneralParamsTable = Build.Write36213_Customers(customerGeneralParamsFile, customerTable);
			
			
			buttonOpenCustomerBusinessRelationFile.Enabled = true;
			buttonOpenCustomerFinancialFile.Enabled = true;
			buttonOpenCustomerFile.Enabled = true;
			buttonOpenCustomerDeliveryFile.Enabled = true;
			buttonOpenCustomerTreeFile.Enabled = true;
			buttonOpenCustomerItemFile.Enabled = true;
			buttonOpenCustomerGeneralParamsFile.Enabled = true;
			buttonCustomer.Enabled = true;
		}
		
		void ButtonOpenSupplierBusinessRelationFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierBusinessRelationFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierBusinessRelationFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierBusinessRelationFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenSupplierFinancialFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierFinancialFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierFinancialFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierFinancialFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenSupplierItemFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierItemFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierItemFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierItemFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenSupplierV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonSupplierClick(object sender, EventArgs e)
		{
			buttonOpenSupplierBusinessRelationFile.Enabled = false;
			buttonOpenSupplierFinancialFile.Enabled = false;
			buttonOpenSupplierFile.Enabled = false;
			buttonOpenSupplierItemFile.Enabled = false;
			buttonOpenSupplierV9File.Enabled = false;
			buttonOpenSupplierV9ItemFile.Enabled = false;
			buttonOpenSupplierGeneralParamsFile.Enabled = false;
			buttonOpenSupplierCodeV9QAD2008File.Enabled = false;
			buttonSupplier.Enabled = false;
			
			System.Data.DataTable businessRelationFile = Build.GetDataTableFromCsvFile(textBoxSupplierBusinessRelationFile.Text);
			System.Data.DataTable supplierFinancialFile = Build.GetDataTableFromCsvFile(textBoxSupplierFinancialFile.Text);
			System.Data.DataTable supplierFile = Build.GetDataTableFromCsvFile(textBoxSupplierFile.Text);
			System.Data.DataTable supplierItemFile = Build.GetDataTableFromCsvFile(textBoxSupplierItemFile.Text);
			System.Data.DataTable supplierV9File = Build.GetDataTableFromCsvFile(textBoxSupplierV9File.Text);
			System.Data.DataTable supplierV9ItemFile = Build.GetDataTableFromCsvFile(textBoxSupplierV9ItemFile.Text);
			System.Data.DataTable supplierCodeV9QAD2008File = Build.GetDataTableFromCsvFile(textBoxSupplierCodeV9QAD2008File.Text);
			System.Data.DataTable supplierGeneralParamsFile = Build.GetDataTableFromCsvFile(textBoxSupplierGeneralParamsFile.Text);
			
			System.Data.DataTable businessRelationTable = Build.Write361431_Suppliers(businessRelationFile, supplierV9File, supplierCodeV9QAD2008File);
			System.Data.DataTable supplierFinancialTable = Build.Write282011_Suppliers(supplierFinancialFile, supplierV9File, supplierCodeV9QAD2008File);
			
			/* TEMPORAIRE */
			System.Data.DataTable supplierTable = Build.Write231_Suppliers(supplierFinancialFile, supplierV9File, supplierCodeV9QAD2008File);
			/* FIN */
			
			System.Data.DataTable supplierGeneralParamsTable = Build.Write36213_Suppliers(supplierGeneralParamsFile, supplierFile, supplierV9File);
			
			buttonOpenSupplierBusinessRelationFile.Enabled = true;
			buttonOpenSupplierFinancialFile.Enabled = true;
			buttonOpenSupplierFile.Enabled = true;
			buttonOpenSupplierItemFile.Enabled = true;
			buttonOpenSupplierV9File.Enabled = true;
			buttonOpenSupplierV9ItemFile.Enabled = true;
			buttonOpenSupplierGeneralParamsFile.Enabled = true;
			buttonOpenSupplierCodeV9QAD2008File.Enabled = true;
			buttonSupplier.Enabled = true;
		}
		
		void ButtonOpenRoutingFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxRoutingFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxRoutingFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxRoutingFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateRouting();
		}
		
		void ButtonOpenRoutingV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxRoutingV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxRoutingV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxRoutingV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateRouting();
		}
		
		void ButtonRoutingClick(object sender, EventArgs e)
		{
			buttonOpenRoutingFile.Enabled = false;
			buttonOpenRoutingV9File.Enabled = false;
			buttonOpenRoutingCommentFile.Enabled = false;
			buttonOpenRoutingCommentV9File.Enabled = false;
			buttonRouting.Enabled = false;
			
			System.Data.DataTable itemFile = Build.GetDataTableFromCsvFile(textBoxItemFile.Text);
			System.Data.DataTable itemSiteCellProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemSiteCellProdLineFile.Text);
			System.Data.DataTable itemProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemProdLineFile.Text);
			System.Data.DataTable itemLeaderFile = Build.GetDataTableFromCsvFile(textBoxItemLeaderFile.Text);
			System.Data.DataTable itemRawFile = Build.GetDataTableFromCsvFile(textBoxItemRawFile.Text);
			System.Data.DataTable itemV9File = Build.GetDataTableFromCsvFile(textBoxItemV9File.Text);
			System.Data.DataTable itemV9ProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9ProdLineFile.Text);
			System.Data.DataTable itemV9LastProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9LastProdLineFile.Text);
			System.Data.DataTable itemDSRPFile = Build.GetDataTableFromCsvFile(textBoxItemDSRPFile.Text);
			
			System.Data.DataTable routingFile = Build.GetDataTableFromCsvFile(textBoxRoutingFile.Text);
			System.Data.DataTable routingV9File = Build.GetDataTableFromCsvFile(textBoxRoutingV9File.Text);
			System.Data.DataTable routingCommentFile = Build.GetDataTableFromCsvFile(textBoxRoutingCommentFile.Text);
			System.Data.DataTable routingCommentV9File = Build.GetDataTableFromCsvFile(textBoxRoutingCommentV9File.Text);
			
			
			System.Data.DataTable itemTable = Build.Write141_Items(itemFile, 
					                                               itemSiteCellProdLineFile,
					                                               itemProdLineFile,
												                   itemLeaderFile, 
												                   itemRawFile,
												                   itemV9File,
												                   itemV9ProdLineFile, 
												                   itemV9LastProdLineFile,
												                   itemDSRPFile);
			
			
			if (itemTable != null)
			{
				System.Data.DataTable routingTable = Build.Write14131_Routing(routingFile, routingV9File, routingCommentFile, routingCommentV9File, itemTable);
			}
			
			buttonOpenRoutingFile.Enabled = true;
			buttonOpenRoutingV9File.Enabled = true;
			buttonOpenRoutingCommentFile.Enabled = true;
			buttonOpenRoutingCommentV9File.Enabled = true;
			buttonRouting.Enabled = true;
		}
		
		void ButtonOpenCustomerGeneralParamsFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxCustomerGeneralParamsFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxCustomerGeneralParamsFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxCustomerGeneralParamsFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateCustomer();
		}
		
		void ButtonOpenSupplierGeneralParamsFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierGeneralParamsFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierGeneralParamsFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierGeneralParamsFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonProdStructClick(object sender, EventArgs e)
		{
			buttonOpenProdStructCodeFile.Enabled = false;
			buttonOpenProdStructFile.Enabled = false;
			buttonOpenProdStructCodeV9File.Enabled = false;
			buttonOpenProdStructV9File.Enabled = false;
			buttonProdStruct.Enabled = false;
			
			System.Data.DataTable itemFile = Build.GetDataTableFromCsvFile(textBoxItemFile.Text);
			System.Data.DataTable itemSiteCellProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemSiteCellProdLineFile.Text);
			System.Data.DataTable itemProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemProdLineFile.Text);
			System.Data.DataTable itemLeaderFile = Build.GetDataTableFromCsvFile(textBoxItemLeaderFile.Text);
			System.Data.DataTable itemRawFile = Build.GetDataTableFromCsvFile(textBoxItemRawFile.Text);
			System.Data.DataTable itemV9File = Build.GetDataTableFromCsvFile(textBoxItemV9File.Text);
			System.Data.DataTable itemV9ProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9ProdLineFile.Text);
			System.Data.DataTable itemV9LastProdLineFile = Build.GetDataTableFromCsvFile(textBoxItemV9LastProdLineFile.Text);
			System.Data.DataTable itemDSRPFile = Build.GetDataTableFromCsvFile(textBoxItemDSRPFile.Text);
			
			System.Data.DataTable prodStructCodeFile = Build.GetDataTableFromCsvFile(textBoxProdStructCodeFile.Text);
			System.Data.DataTable prodStructCodeV9File = Build.GetDataTableFromCsvFile(textBoxProdStructCodeV9File.Text);
			System.Data.DataTable prodStructFile = Build.GetDataTableFromCsvFile(textBoxProdStructFile.Text);
			System.Data.DataTable prodStructV9File = Build.GetDataTableFromCsvFile(textBoxProdStructV9File.Text);
			
			
			System.Data.DataTable itemTable = Build.Write141_Items(itemFile, 
					                                               itemSiteCellProdLineFile,
					                                               itemProdLineFile,
												                   itemLeaderFile, 
												                   itemRawFile,
												                   itemV9File,
												                   itemV9ProdLineFile, 
												                   itemV9LastProdLineFile,
												                   itemDSRPFile);
			
			if (itemTable != null)
			{
				System.Data.DataTable prodStructCodeTable = Build.Write131_CodeProdStruct(prodStructCodeFile, prodStructCodeV9File, itemTable);
				System.Data.DataTable prodStructTable = Build.Write135_ProdStruct(prodStructFile, prodStructV9File, itemTable);
			}
			
			buttonOpenProdStructCodeFile.Enabled = true;
			buttonOpenProdStructFile.Enabled = true;
			buttonOpenProdStructCodeV9File.Enabled = true;
			buttonOpenProdStructV9File.Enabled = true;
			buttonProdStruct.Enabled = true;
		}
		
		void ButtonOpenProdStructCodeV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxProdStructCodeV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxProdStructCodeV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxProdStructCodeV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateProdStruct();
		}
		
		void ButtonOpenProdStructCodeFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxProdStructCodeFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxProdStructCodeFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxProdStructCodeFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateProdStruct();
		}
		
		void ButtonOpenProdStructFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxProdStructFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxProdStructFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxProdStructFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateProdStruct();
		}
		
		void ButtonOpenProdStructV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxProdStructV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxProdStructV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxProdStructV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateProdStruct();
		}
		
		void ButtonOpenItemCostFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemCostFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemCostFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemCostFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemV9CostFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemV9CostFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemV9CostFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemV9CostFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenSupplierFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenSupplierV9ItemFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierV9ItemFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierV9ItemFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierV9ItemFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenItemIntrastatCodeFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemIntrastatCodeFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemIntrastatCodeFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemIntrastatCodeFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemAnalysisCodeBrandFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemAnalysisCodeBrandFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemAnalysisCodeBrandFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemAnalysisCodeBrandFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenWorkCenterFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxWorkCenterFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxWorkCenterFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxWorkCenterFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateWorkCenter();
		}
		
		void ButtonOpenWorkCenterV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxWorkCenterV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxWorkCenterV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxWorkCenterV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateWorkCenter();
		}
		
		void ButtonWorkCenterClick(object sender, EventArgs e)
		{
			buttonOpenWorkCenterFile.Enabled = false;
			buttonOpenWorkCenterV9File.Enabled = false;
			buttonWorkCenter.Enabled = false;
			
			System.Data.DataTable workCenterFile = Build.GetDataTableFromCsvFile(textBoxWorkCenterFile.Text);
			System.Data.DataTable workCenterV9File = Build.GetDataTableFromCsvFile(textBoxWorkCenterV9File.Text);
			
			System.Data.DataTable workCenterTable = Build.Write145_WorkCenter(workCenterFile, workCenterV9File);
			
			buttonOpenWorkCenterFile.Enabled = true;
			buttonOpenWorkCenterV9File.Enabled = true;
			buttonWorkCenter.Enabled = true;
		}
		
		void ButtonOpenProductionLineFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxProductionLineFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxProductionLineFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxProductionLineFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateProductionLine();
		}
		
		void ButtonProductionLineClick(object sender, EventArgs e)
		{
			buttonOpenProductionLineFile.Enabled = false;
			buttonProductionLine.Enabled = false;
			
			System.Data.DataTable productionLineFile = Build.GetDataTableFromCsvFile(textBoxProductionLineFile.Text);
			
			System.Data.DataTable productionLineTable = Build.Write182211_ProductionLine(productionLineFile);
			
			buttonOpenProductionLineFile.Enabled = true;
			buttonProductionLine.Enabled = true;
		}
		
		void ButtonOpenRoutingCommentFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxRoutingCommentFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxRoutingCommentFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxRoutingCommentFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateRouting();
		}
		
		void ButtonOpenRoutingCommentV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxRoutingCommentV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxRoutingCommentV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxRoutingCommentV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateRouting();
		}
		
		void ButtonOpenItemCommentFCFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemCommentFCFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemCommentFCFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemCommentFCFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemCommentFBFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemCommentFBFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemCommentFBFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemCommentFBFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemCommentWAFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemCommentWAFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemCommentWAFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemCommentWAFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemLogisticsFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemLogisticsFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemLogisticsFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemLogisticsFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenItemV9LogisticsFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxItemV9LogisticsFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxItemV9LogisticsFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxItemV9LogisticsFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateItem();
		}
		
		void ButtonOpenMeasureFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxMeasureFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxMeasureFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxMeasureFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateMeasure();
		}
		
		void ButtonOpenMeasureV9FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxMeasureV9File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxMeasureV9File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxMeasureV9File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateMeasure();
		}
		
		void ButtonMeasureClick(object sender, EventArgs e)
		{
			buttonOpenMeasureFile.Enabled = false;
			buttonOpenMeasureV9File.Enabled = false;
			
			System.Data.DataTable measureFile = Build.GetDataTableFromCsvFile(textBoxMeasureFile.Text);
			System.Data.DataTable measureV9File = Build.GetDataTableFromCsvFile(textBoxMeasureV9File.Text);
			
			System.Data.DataTable measureTable = Build.Write113_Measure(measureFile, measureV9File);
			
			buttonOpenMeasureFile.Enabled = true;
			buttonOpenMeasureV9File.Enabled = true;
		}
		
		private void UpdateMeasure()
		{
			if ((textBoxMeasureFile.Text != "") &&
			    (textBoxMeasureV9File.Text != ""))
			{
				buttonMeasure.Enabled = true;
			}
			else
			{
				buttonMeasure.Enabled = false;
			}
		}
		
		void ButtonOpenSupplierCodeV9QAD2008FileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierCodeV9QAD2008File.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierCodeV9QAD2008File.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierCodeV9QAD2008File.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonOpenSupplierPricingFileClick(object sender, EventArgs e)
		{
			openFileDialogCsv = new OpenFileDialog();
			openFileDialogCsv.Multiselect = false;
			openFileDialogCsv.Filter = "CSV files (*.csv)|*.csv";
			
			if (textBoxSupplierPricingFile.Text != "")
			{
				openFileDialogCsv.FileName = textBoxSupplierPricingFile.Text;
			}
			
			if (openFileDialogCsv.ShowDialog() == DialogResult.OK)
			{
				textBoxSupplierPricingFile.Text = openFileDialogCsv.FileName;
			}
			
			UpdateSupplier();
		}
		
		void ButtonSupplierPricingClick(object sender, EventArgs e)
		{
			
		}
	}
}
