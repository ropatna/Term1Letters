using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class Centre_intimation : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=CBSEPAT\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public Centre_intimation()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            string database = Microsoft.VisualBasic.Interaction.InputBox("ENTER NAME OF DATABASE FROM WHICH LETTER HAS TO BE GENERATED", "INPUT DATABASE NAME", "bk2021");
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where cen_schno='" + textBox1.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Document pdoc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(pdoc, new FileStream("F:\\pdf\\centre\\" + dr["cen_schno"].ToString() + "_" + dr["cenno"].ToString() + ".pdf", FileMode.Create));
                var header = iTextSharp.text.Image.GetInstance("D:\\WindowsFormsApp1\\WindowsFormsApp1\\images\\header.png");
                var footer = iTextSharp.text.Image.GetInstance("D:\\WindowsFormsApp1\\WindowsFormsApp1\\images\\FOOTER.png");
                var rosign = iTextSharp.text.Image.GetInstance("D:\\WindowsFormsApp1\\WindowsFormsApp1\\images\\rosignpng.png");
                var header2 = iTextSharp.text.Image.GetInstance("D:\\WindowsFormsApp1\\WindowsFormsApp1\\images\\acceptance_header.png");
                header2.ScaleToFit(900f, 60f);
                header2.Alignment = 1;
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(90f, 30f);
                rosign.SetAbsolutePosition(480, 360);
                pdoc.Open();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                pdoc.AddTitle("Centre Intimation Letter");
                //
                Paragraph p = new Paragraph("===============================================================================\n");
                pdoc.Add(p);
                Paragraph p2 = new Paragraph("CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p2);
                Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p3);
                Paragraph p4 = new Paragraph(str: "F.NO.RO/PTN/CONF/MAIN/2021/      (" + dr["cen_schno"].ToString() + ")                                                Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p4);
                //Paragraph p50 = new Paragraph("REVISED", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                //pdoc.Add(p50);
                Paragraph p5 = new Paragraph(str: "CENTER NO. " + dr["cenno"].ToString() + "\n                      ======") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p5);
                Paragraph p6 = new Paragraph(str: "The Principal,(" + dr["cen_schno"].ToString() + ")\n" + dr["add1"].ToString() + "\n" + dr["add2"].ToString() + "\n" + dr["add3"].ToString() + "\n" + dr["add4"].ToString() + "\n" + dr["add5"].ToString() + "\n" + dr["pin"].ToString() + "\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p6);
                Paragraph p7 = new Paragraph(str: "Sub:  Intimation & Acceptance proforma regarding fixation of Examination centre\n                        for AISSE / AISSCE - 2021(MAIN).") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p7);
                Paragraph p8 = new Paragraph(str: "**********************************************************************************************************************\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p8);
                Paragraph p9 = new Paragraph(str: "Dear Principal,") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p9);
                Paragraph p10 = new Paragraph(str: "     I am glad to inform you that your School has been  fixed as an Examination Centre for the ensuing MAIN Examination of AISSE/AISSCE-2021. The Board Examinations is Scheduled to be commence from 4th May 2021. ( Datesheet available on Board's Website, i.e., https:www.cbse.nic.in )\n\n      In this regard, I am to draw your kind attention that it is obligatory on the part of the Principals of CBSE affiliated Schools to act  as  a  Centre  Superintendent and  appoint elligible teachers  as  invigilators  at  the Examination Centre.  Also as per affiliation Bye-laws, you need to keep the  School  Building Staff and other equipment at the disposal of the Board for conducting the  Examination.  The Total number of candidates (approximate)  allotted  at  your  Centre are as follows:\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p10);
                /*Centre's No. of Candidates*/
                PdfPTable tbl = new PdfPTable(2);
                tbl.HorizontalAlignment = 2;
                tbl.DefaultCell.Border = 0;
                tbl.AddCell(new Phrase("AISSE-Main(CLASS-X)"));
                tbl.AddCell(new Phrase(dr["tot10"].ToString()));
                tbl.AddCell(new Phrase("AISSCE(CLASS - XII)"));
                tbl.AddCell(new Phrase(dr["tot12"].ToString()));
                pdoc.Add(tbl);
                //
                Paragraph p11 = new Paragraph(str: "      Question Papers of various Subjects pertaining to your Centre shall be stored at the following bank as per the details furnished by you which  may please be kept STRICTLY CONFIDENTIAL.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p11);
                /*Custodian Details*/
                PdfPTable tbl2 = new PdfPTable(2);
                tbl2.HorizontalAlignment = 1;
                tbl2.AddCell(new Phrase(dr["cust_name"].ToString() + "\n" + dr["BKADD1"].ToString() + "\n" + dr["BKADD2"].ToString() + "\n" + dr["BKADD3"].ToString() + "\n" + dr["bkdistt"].ToString() + "\n" + dr["bkstate"].ToString() + "\nPIN:" + dr["bkpin"].ToString()));
                tbl2.AddCell(new Phrase("Bank Contact No.\n" + dr["bkph1"].ToString() + " / " + dr["bkph2"].ToString() + "\n\nBANK MANAGER NAME & CONTACT:\n" + dr["bmname"].ToString() + "\n" + dr["bmph1"].ToString() + " / " + dr["bmph2"].ToString()));
                pdoc.Add(tbl2);
                //
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p12 = new Paragraph(str: "       The Question papers may be verified well in  advance  after  receipt in  the Bank without disturbing the Seals / Packets so as to ensure that Question papers of Subjects for which Students are appearing at your Centre are intact and available and the same may be collected from the said bank  well  in  time on the  days  of Examination so as to reach the Examination Centre at 9:30 A.M.positively as per the Guidelines provided  by  the  Board  for  Centre Superintendents.  In  case, your School Centre is considered a  SELF CENTRE, the  status  of  appointment of Centre Superintendent will be informed separately.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p12);
                Paragraph p13 = new Paragraph(str: "       Further, I am also to inform you that in case your School has been fixed  as a Self Examination Centre and does not receive intimation about  the  appointment of Centre Superintendant at your School Centre latest by " + date_str + ", please inform the undersigned immediately by email  on   ropatna.cbse@nic.in or abcell.cbseropatna@gmail.com for necessary action.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p13);
                Paragraph p14 = new Paragraph(str: "       During the conduct of the  Examinations,  you need  to  play  a  significant role in ensuring smooth and fair conduct without giving any scope to  others  for comments and making  necessary  arrangements  for   comfortable seating  of  the Students besides  ensuring  availability of required  furniture, lights, water, ventilation, toilets etc.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p14);
                Paragraph pn1 = new Paragraph(str: "       The school / centre  must  ensure  Social Distancing of candidates at their centre in  light  of  guidelines prescribed by govt.of India for COVID - 19 pandamic. Guidelines to prevent COVID-19 pandamic during Examination of AISSE/AISSCE-2021 shall be made available to the centres in CS Guidelines-2021.\n", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(pn1);
                Paragraph pn2 = new Paragraph(str: "       Both the ALL  India  Sr.Sch Certificate  MAIN Examination(Class XII) and Secondary  School Examination MAIN(Class X)  will be held at the  Examination Centre  and only one  Centre  Superintendent  will  be appointed for the purpose. However, in case the number of candidates allotted is more than 250, one  Deputy Centre Superintendent can be appointed as per Guidelines provided to  the  Centre Superintendants(Copy of Guidelines Shall be supplied to you  along  with  Centre Materials).No TA / DA will be paid to local Centre Superintendents.Second  clerks at the rate of same remuneration may appointed on such sessions when  the No.of candidates exceed 250.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(pn2);
                Paragraph p15 = new Paragraph(str: "       REMUNERATION\\CONVEYANCE IS ADMISSIBLE AS PER LAST YEAR OR AS PER CS GUIDELINES:\n        =================================================================================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 10)) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p15);
                PdfPTable tbl3 = new PdfPTable(2);
                tbl3.HorizontalAlignment = 0;
                tbl3.DefaultCell.Border = 0;
                tbl3.AddCell(new Phrase("CENTRE SUPERINTENDENT"));
                tbl3.AddCell(new Phrase(":  Remuneration @Rs. 350/- per day and Conveyance  @Rs. 250/- per day for local Supdtt. @Rs. 250/- per day +TA/DA for outside C.S."));
                tbl3.AddCell(new Phrase("DEPUTY CENTRE SUPERINTENDENT"));
                tbl3.AddCell(new Phrase(":  Remuneration @Rs. 250/- per day and conveyance  @Rs. 250/- per day."));
                tbl3.AddCell(new Phrase("ASSTT. CENTRE SUPERINTENDENT"));
                tbl3.AddCell(new Phrase(":  Remuneration @Rs. 200/- per day and conveyance  @Rs. 150/- per day"));
                tbl3.AddCell(new Phrase("CLERK"));
                tbl3.AddCell(new Phrase(":  @Rs. 200 per day."));
                tbl3.AddCell(new Phrase("CLASS IV EMPLOYEES"));
                tbl3.AddCell(new Phrase(":  @Rs. 100/- per day."));
                tbl3.AddCell(new Phrase("REFRESHMENT TO THE STAFF OF CENTRE"));
                tbl3.AddCell(new Phrase(":  @RS. 75 per day per head"));
                tbl3.AddCell(new Phrase("PAYMENT TO SCRIBE WITH DISABLED CANDIDATE"));
                tbl3.AddCell(new Phrase(":  @Rs. 150/- per session of Examination through Examination Centre and shall be included in the Centre charges bill."));
                pdoc.Add(tbl3);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p16 = new Paragraph(str: "        ASSTT. SUPERINTENDENT AND CLERKS:", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p16);
                PdfPTable tbl4 = new PdfPTable(2);
                tbl4.HorizontalAlignment = 0;
                tbl4.DefaultCell.Border = 0;
                tbl4.AddCell(new Phrase("(A) In a hall or big rooms"));
                tbl4.AddCell(new Phrase("--  One Asstt. Supdt. for every 12 candidates on part thereof."));
                tbl4.AddCell(new Phrase("(B) In smaller rooms having upto 40 candidates"));
                tbl4.AddCell(new Phrase("--  Two Asstt. Supdts. in each room if the No. of candidates exceed 20."));
                tbl4.AddCell(new Phrase("(C) Clerk"));
                tbl4.AddCell(new Phrase("--  One for each centre."));
                pdoc.Add(tbl4);
                Paragraph p17 = new Paragraph(str: "        CLASS IV EMPLOYEES\n        ====================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p17);
                PdfPTable tbl5 = new PdfPTable(2);
                tbl5.HorizontalAlignment = 0;
                tbl5.DefaultCell.Border = 0;
                tbl5.AddCell(new Phrase("Upto 20 candidates"));
                tbl5.AddCell(new Phrase("--  One"));
                tbl5.AddCell(new Phrase("between 21 to 100 Candidates"));
                tbl5.AddCell(new Phrase("--  Two"));
                tbl5.AddCell(new Phrase("between 101 to 400 Candidates"));
                tbl5.AddCell(new Phrase("--  Three"));
                tbl5.AddCell(new Phrase("401 or More candidates"));
                tbl5.AddCell(new Phrase("--  Four"));
                pdoc.Add(tbl5);
                Paragraph p18 = new Paragraph(str: "      To ensure proper packing of the Answer Books intact and its dispatch on  the same  day  by  speed post / or as per instant directions / by hand(for local Center) Clerk and Class IV  employees may be appointed from the school  itself as far as possible.Persons from Outside may only be appointed in case the Principal  of the school is not in a Position to provide the clerk and the required number of Class IV employees.\n      THE ASSTT. SUPDTS.  SHOULD BE APPOINTED AMONGST  THE  TEACHING STAFF   OF THE SCHOOL AND SHOULD PREFERABLY BE PGT / TGTS WHO  ARE RELIABLE. \"YOU  SHOULD   TAKE AN UNDERTAKING FROM ALL THE STAFF APPOINTED  FOR THE EXAMINATION THAT NO   NEAR RELATION OF HIS / HER IS APPEARING AT THE  EXAMINATION  CONCERNED.\" PERSONS LIVING AT FAR OFF PLACES SHOULD BE DISCOURGED  FOR APPOINTMENT AS ASSTT.SUPDT.AND NO TA / DA IS ADMISSIBLE TO ASSTT.SUPDT.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p18);
                Paragraph p19 = new Paragraph(str: "      A sum of Rs. 10 for infrastructure usage charges including sitting arrangement is payable per candidate on a maximum No.of candidates  appearing  in the  Examination in a day for the entire period of Examinations  and not per session.This is excluding of Rs. 15 per candidate on account of stationery, packing  materials etc. but does not include conveyance charges for  depositing / dispatch  of Answer  Books and postage charges of the Parcel.Rs. 3 per  candidate alloted  to the centre for whole examination towards printing centre material, attendance sheet etc.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p19);
                Paragraph p20 = new Paragraph(str: "      Answer Books,  Supplementary Answer  Books  and    other  related  Examination materials are already available with you and if not  available; as per  requirement, please contact to this office immediately. If an external Centre Superintendant has been appointed at your School Centre, these materials should  be handed over to the Centre Superintendent of your Examination Centre, when he/ she  comes to your School one or  two days before commencement of the Examinations.After the Examinations, unused Answer Books etc. may please be collected  from him/ her with proper accounts and  should be returned to the Board.  The bill  of centre  charges may please be sent immediately after completion of Examination  for  adjustment of  the advance released for 2021  Examinations  with  an undertaking that all payment made through ECS /as per rules.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p20);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p21 = new Paragraph(str: "      The  Guidelines  for  Centre  Superintendent may  thoroughly  be  perused  and followed  meticulously  with  regard  to instructions for  conduct of  Examination and  payment  of  Centre  Charges etc., so as to ensure  smooth  and  fair  conduct of Examinations.\n\n   Please  send  the  ACCEPTANCE  of the  Centre  at your School in the  enclosed proforma  duly  completed  in  all respects by return of  post as well as by email on ropatna.cbse@nic.in/abcell.cbseropatna@gmail.com so as to reach the undersigned latest by " + date_str + "  without fail Positively.\n\n     I am hopeful that you would exercise due care and concern in making necessary arrangements as per Board's rule and extend full co-operation  in  smooth and fair conduct of Secondary School Examination and All India Senior / Secondary School MAIN Examination - 2021\n\n     The duly  filled up  centre charges bills along with proper receipts must be submitted within 30 days of  completion of  work failing which penal interest may be charged by the Board.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p21);
                Paragraph p22 = new Paragraph(str: "Wishing  you  successful, smooth and fair conduct of Main Examinations-2021.\n\nWith BEST WISHES.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p22);
                Paragraph p23 = new Paragraph(str: "Yours faithfully,    ") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p23);
                pdoc.Add(rosign);
                Paragraph p24 = new Paragraph(str: "\n(J BARMAN)        \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p24);
                pdoc.NewPage();
                pdoc.Add(header2);
                pdoc.Add(footer);
                Paragraph p25 = new Paragraph(str: "CENTRE ACCEPTANCE\n===================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p25);
                Paragraph p26 = new Paragraph(str: "CENTRE NO.  " + dr["cenno"].ToString() + "       \n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p26);
                PdfPTable tbl6 = new PdfPTable(4);
                tbl6.SetWidths(new float[] { 100f, 2f, 30f, 2f });
                tbl6.HorizontalAlignment = 0;
                tbl6.DefaultCell.Border = 0;
                tbl6.AddCell(new Phrase("\nThe Regional Officer,\nCentral Board of Secondary Education,\nAmbika Complex,Behind State Bank Colony\nNear Brahmsthan,Sheikhpura, Bailey Road\nPatna (Bihar) - 800 014\n"));
                tbl6.AddCell(new Phrase(":\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:"));
                tbl6.AddCell(new Phrase("-----------------------\nPlease affix Passport Size Photograph of Centre Supritendent and do the full signature bellow.\n-----------------------\n\n\n\n-----------------------"));
                tbl6.AddCell(new Phrase(":\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:\n:"));
                pdoc.Add(tbl6);
                Paragraph p27 = new Paragraph(str: "Sir,\n       With reference to your letter No. F.NO.RO/PTN/CONF/MAIN/2021/     / " + dr["cen_schno"].ToString() + " / " + date_str + ". I hereby accept our School as an Examination Centre for  Conducting the AISSE / AISSCE MAIN 2021 Examination(Cen.No " + dr["cenno"].ToString() + " ) as per the instructions Guidelines issued by the Board time to time and all efforts would be made to ensure smooth and fair conduct of the AISSE / AISSCE 2021(MAIN).\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p27);
                /*Bank Details and Custodian's Details*/
                PdfPTable tbl7 = new PdfPTable(2);
                tbl7.WidthPercentage = 100f;
                tbl7.HorizontalAlignment = 1;
                tbl7.AddCell(new Phrase("Details of School\n=================="));
                tbl7.AddCell(new Phrase("Details of Bank Custodian\n=========================="));
                tbl7.AddCell(new Phrase(" " + dr["add1"].ToString() + "\n" + dr["add2"].ToString() + "\n" + dr["add3"].ToString() + "\n" + dr["add4"].ToString() + "\n" + dr["add5"].ToString() + "\nPIN:" + dr["pin"].ToString() + "\nBANK NAME: " + dr["bnkname"].ToString() + "\nACCOUNT NO.: " + dr["bnkaccount"].ToString() + "\nIFSC CODE: " + dr["bnkifsc"].ToString() + "\nName & Contact No.:\n" + dr["prname"].ToString() + "\n" + dr["pmob1"].ToString()));
                //tbl7.AddCell(new Phrase(" " + dr["add1"].ToString() + "\n" + dr["add2"].ToString() + "\n" + dr["add3"].ToString() + "\n" + dr["add4"].ToString() + "\n" + dr["add5"].ToString() + "\nPIN:" + dr["pin"].ToString() + "\nBANK NAME: " + dr["bnkname"].ToString() + "\nACCOUNT NO.: " + dr["bnkaccount"].ToString() + "\nIFSC CODE: " + dr["bnkifsc"].ToString() + "\nName & Contact No.:\n" + dr["prname"].ToString() + "\n" + dr["pmob1"].ToString() + "      " + dr["pmob2"].ToString() + "\n" + dr["pmobile"].ToString() + "      " + dr["pmobile2"].ToString() + "\n"));
                tbl7.AddCell(new Phrase(" " + dr["cust_name"].ToString() + "\n" + dr["bkadd1"].ToString() + "\n" + dr["bkadd2"].ToString() + "\n" + dr["bkadd3"].ToString() + "\n" + dr["bkdistt"].ToString() + "     " + dr["bkstate"].ToString() + "\nBranch Code: " + dr["brcode"].ToString() + "\nCUSTODIAN ACCOUNT NO.:" + dr["custaccnt"].ToString() + "\nCUSTODIAN IFSC CODE: " + dr["custifsc"].ToString() + "\nName & Contact No.:\n" + dr["bmname"].ToString() + "\n" + dr["bmph1"].ToString() + "    (This Number will be used by Branch Manager for CMTM-Mobile app for Custodians.) \n" + dr["bmph2"].ToString() + "\n"));
                pdoc.Add(tbl7);
                //
                Paragraph p30 = new Paragraph(str: "(NOTE: IFSC Code has been changed due to merger of Bank, Kindly make correction on the above and report to Regional Office Patna.)\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p30);
                Paragraph p28 = new Paragraph(str: "I undertake that information  mentioned above are true and there is no change in above records.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p28);
                Paragraph p29 = new Paragraph(str: "Yours faithfully,    \n\nSignature:__________________\n" + dr["prname"].ToString() + "\nDesignation:  " + dr["post"].ToString() + "            \nLatest Mobile No. for using CMTM mobile app by CS..............................\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p29);
                pdoc.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
