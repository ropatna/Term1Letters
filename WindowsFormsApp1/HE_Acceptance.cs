﻿using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class HE_Acceptance : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        string d1 = ""; //STARTING DATE OF EXAM
        string d2 = ""; //ENDING DATE OF EXAM
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=CBSEPAT\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public HE_Acceptance()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            string database = Microsoft.VisualBasic.Interaction.InputBox("ENTER NAME OF DATABASE FROM WHICH LETTER HAS TO BE GENERATED", "INPUT DATABASE NAME", "HEPAT2021");
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where HENO='" + textBox2.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Document pdoc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(pdoc, new FileStream("F:\\pdf\\he\\" + dr["he_sch"].ToString() + "_" + dr["heno"].ToString() + ".pdf", FileMode.Create));
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
                rosign.SetAbsolutePosition(470, 135);
                pdoc.Open();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                pdoc.AddTitle("HE Acceptance Letter");
                //
                Paragraph p = new Paragraph("===============================================================================\n");
                pdoc.Add(p);
                Paragraph p2 = new Paragraph("CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p2);
                Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p3);
                Paragraph p4 = new Paragraph(str: "Ref.No.:RO(PTN)/CONF./H.E./" + dr["hesub"].ToString() + "/" + dr["heno"].ToString() + "/2021-MAIN/                                                   Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p4);
                Paragraph p5 = new Paragraph(str: dr["he_sch"].ToString() + "\n" + dr["hename"].ToString() + " , " + dr["he_desig"].ToString() + "\n" + dr["headd1"].ToString() + "\n" + dr["headd2"].ToString() + "\n" + dr["headd3"].ToString() + "\n" + dr["headd4"].ToString() + "\n" + dr["headd5"].ToString() + "\nPIN: " + dr["hepin"].ToString() + "\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p5);
                Paragraph p6 = new Paragraph(str: "Sub:  APPOINTMENT LETTER OF HEAD EXAMINER FOR CLASS " + dr["HECLASS"].ToString() + " SUBJECT (" + dr["hesub"].ToString() + dr["hesub2"].ToString() + ")\n          (" + dr["subname"].ToString() + dr["subname2"].ToString() + ") EXAMINATION 2021.\n\nSir/Madam,\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p6);
                if (dr["HECLASS"].ToString() == "10")
                {
                    d1 = "04-05-2021";
                    d2 = "07-06-2021";
                }
                if (dr["HECLASS"].ToString() == "12")
                {
                    d1 = "04-05-2021";
                    d2 = "14-06-2021";
                }
                Paragraph p7 = new Paragraph(str: "      The Board is pleased  to  appoint  you  as Head  Examiner  for the  subject  " + dr["subname"].ToString() + dr["subname2"].ToString() + "(" + dr["hesub"].ToString() + dr["hesub2"].ToString() + ") of Class " + dr["HECLASS"].ToString() + " Examination to be held from " + d1 + " to " + d2 + ". \n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p7);
                Paragraph p8 = new Paragraph(str: "      In order to ensure proper manageability, effective supervision and timely completion of evaluation work with accuracy, the Board is also appointing Chief Nodal Supervisor at selected cities.  The Chief Nodal Supervisor would be a Principal of the School and under his / her supervision 3 to 10 Vice Principals / PGT(s) will be appointed as Head  Examiner.  The   Head  Examiner  would  do the evaluation work in the school of Chief  Nodal Supervisor  and the  time of evaluation will  be decided  by you  in  consultation with the Chief Nodal Supervisor as per schedule.  The tentative  date  for  commencement of  Spot Evaluation in the  subject is after  05 - 06 days  from the date of conduct of Examination.  The Answer Books will be handed over  to  you for  evaluation by Chief Nodal Supervisor.  The details  of your Nodal Centre is given hereunder:-") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p8);
                Paragraph p9 = new Paragraph(str: "      The Evaluation will be done at :\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p9);
                PdfPTable tbl = new PdfPTable(1);
                tbl.HorizontalAlignment = 1;
                tbl.AddCell(new Phrase(dr["cns_schno"].ToString() + "\n" + dr["cnsadd1"].ToString() + "\n" + dr["cnsadd2"].ToString() + "\n" + dr["cnsadd3"].ToString() + "\n" + dr["cnsadd4"].ToString() + "\n" + dr["cnsadd5"].ToString() + "\n" + dr["cnspin"].ToString() + "\nCNS MOBILE NO.: " + dr["cnsmobile"].ToString()));
                pdoc.Add(tbl);
                pdoc.NewPage();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p10 = new Paragraph(str: "      The   appointment  as  well  as  the  information   in  this   regard will  be kept  secret   by  you.  The  Coordinator & Additional  Head   Examiner will be appointed as per  instructions given in the  guidelines  for  Spot or Nodal Evaluation Centres.  THE COORDINATION WORK WILL BE CARRIED  OUT  STRICTLY AS PER INSTRUCTIONS GIVEN IN THE GUIDELINES.  A list of Examiners attached with you will be sent in due course.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p10);
                if ((dr["cls"].ToString() == "10") && dr["hesub"].ToString() == "086")
                {
                    Paragraph c = new Paragraph(str: "      The Evaluation of Science Class X Answer Scripts would be done by two Examiniers. Section - A of the Question Paper will be Evaluated  by the teachers whose subjects are Physics and Chemistry at B.Sc.level and Section - B  of  the Science Question Paper will be Evaluated by the  teachers  whose  subjects is Biology at the B.Sc.The remuneration  for evaluation of Section A and B would be Rs. 18/- and Rs. 7/- respectively.") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(c);
                }
                Paragraph p11 = new Paragraph(str: "      The delivery of Answer Books for evaluation will  be made tentatively after 05 - 06 days from the commencement of the Examination of the concerned subject and evaluation will start immediately  on  receipt  of  Answer  Books  by  the HE / CNS.  The appointment as well as the information in this regard will be kept strictly confidential by you.  AS A HEAD EXAMINER, YOU WILL HAVE TO ENSURE THAT THE EXAMINERS WORKING UNDER YOU, EVALUATE THE ANSWER SCRIPT STRICTLY IN ACCORDANCE WITH THE MARKING SCHEME WHICH LEAVES NO SCOPE FOR ANY  OTHER  CONSIDERATION / ALLEGATION  EXCEPT  THE  REAL  MERIT / PERFORMANCE  OF  THE  CANDIDATE.  It is also expected that the Coordinator  will perform their duties with utmost care  and   sense  of responsibility  and shall  see to  it  that  coordination work done by them is absolutely free of errors.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p11);
                Paragraph p12 = new Paragraph(str: "      FURTHER, AS PER PAST EXPERIENCE IT HAS BEEN OBSERVED THAT QUALITY  OF EVALUATION DONE AT SOME OF SPOT EVALUATION CENTRES WAS  NOT FOUND SATISFACTORY AND ALSO NOT UPTO  THE  DESIRED  LEVEL.  A LARGE NUMBER OF MISTAKE CASES WERE DETECTED  DURING  THE  COURSE  OF  SCRUTINY ON ACCOUNT OF  EVALUATION  OF 2020 EXAMINATION, WHICH  WERE  NOT  DONE  PROPERLY & RAISED  QUESTIONS  ON   THE CREDIBILITY ON WORKING  OF  HEAD  EXAMINERS  AND  EXAMINERS  WHO  PARTICIPATED IN THE EVALUATION WORK.  AS YOU ARE AWARE  THAT  THE  STUDENTS  LOOK  AT  THESE EXAMINATION AS A FINAL EVALUATION OF THEIR ACADEMIC PERFORMANCE.  THE COMPETENT AUTHORITY OF THE BOARD HAS TAKEN IT SERIOUSLY OWING TO LARGE NO. OF  MISTAKES DURING   EVALUATION.  ALSO  FROM  EXAM  2012  THE   STUDENTS / EXAMINEES   MAY TAKE  PHOTOCOPY OF THEIR ANSWER SHEETS UNDER RTI ACT 2005 AS PER ORDERS OF THE HON'BLE SUPREME COURT OF INDIA.  AS  FROM  2014  EXAM  CANDIDATE  CAN  GET  RE-EVALUATED THEIR ANSWER BOOKS ON CERTAIN SUBJECTS.  THEREFORE IT IS REQUESTED THAT PROPER ATTENTION TOWARDS EVALUATION SHOULD BE GIVEN AND ANSWER BOOKS  OF  THE SUBJECT BE EVALUATED IN PERFECT MANNER STRICTLY IN ACCORDANCE WITH THE MARKING SCHEME.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p12);
                pdoc.NewPage();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p13 = new Paragraph(str: "      APART FROM ABOVE, THE EXAMINERS  EVALUATING  THE  ANSWER BOOKS OF THE MEDIUM  OTHER THAN THE ONE THEY ARE TEACHING IN, MAY HAVE  SOME DIFFICULTY  IN UNDERSTANDING  THE  ANSWER   WHICH   MAY   LEAD   TO  WRONG  AWARD  OF  MARKS. THEREFORE, FOR DOING FULL JUSTICE TO  THE  STUDENTS, ALL HEAD EXAMINERS SHOULD STRICTLY  CHECK  AND  ENSURE THAT NO ANSWER BOOK IS EVALUATED BY EXAMINERS WHO ARE TEACHING IN THE MEDIUM OTHER THAN THE ONE USED IN THE ANSWER BOOK i.e.ALL ANSWER BOOKS SHOULD  BE  CHECKED BY THE  SUBJECT EXAMINERS OF THE SAME MEDIUM.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p13);
                Paragraph p14 = new Paragraph(str: "      Further, Head  Examiner  should  check  whether  any  Answer Books of Physically Challanged children(Spastic, Blind, Physically  Handicapped  and Dyslexic children) have been erroneously received alongwith the  Answer  Books of other candidates.  If the Answer Books of Physically Challanged children are found mixed with the Answer Books of other candidates, these  be  immediately returned  to  the undersigned unevaluated through Sealed/Insured Speed  Post Parcel in consultation with the Chief Nodal Supervisor.\n      After  Evaluation,  the  Answer  Books   will   be   serialised    in ascending order and should be packed centurywise in respective  Answer  Books bags.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p14);
                Paragraph p15 = new Paragraph(str: "      The Remunerations, Conveyance and Refreshment charges will be paid to you and sub - examiners as per spot Evaluation Guidelines 2021.\n      It may also be  noted  that  the  Evaluation  is the  most  important part  of the whole Examination system and it determines  the  future career of the  students.  Therefore you are  requested  to  take  every possible care to ensure objective and  judicious  Evaluation  to  safeguard the interest of the students and also to avoid any future complications / allegations.\n      A  copy of the guidelines for Spot Evaluation will be sent to you in due course for your reference and strict compliance.\n      You are requested to send the  bill within one  month  after  completion of work with following details of your Bank A/C:\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p15);
                PdfPTable tbl2 = new PdfPTable(1);
                tbl2.HorizontalAlignment = 1;
                tbl2.DefaultCell.Border = 0;
                tbl2.AddCell(new Phrase("a) A/C holder name\nb) IFSC code\nc) A/C No.\nd) Bank name branch & location\n\n"));
                pdoc.Add(tbl2);
                Paragraph p16 = new Paragraph(str: "      Your acceptance of the assignment being offered by the Board should  reach  to the undersigned immediately in the enclosed proforma  duly  completed  in all  respects through  SPEED POST / EMAIL: abcell.cbseropatna@gmail.com / ropatna.cbse@nic.in .\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p16);
                Paragraph p17 = new Paragraph(str: "      Yours faithfully,\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p17);
                pdoc.Add(rosign);
                Paragraph p18 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p18);
                pdoc.NewPage();
                pdoc.Add(header2); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p19 = new Paragraph(str: "ACCEPTANCE FORM FOR HEAD EXAMINER\nSENIOR / SECONDARY SCHOOL EXAMINATION 2021\n------------------------------------------------------------------------------------------------------------------------------------------\n(To be sent by Speed Post/Email or through Messanger)\nIMMEDIATE & CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p19);
                Paragraph p20 = new Paragraph(str: "Regn.No. " + dr["heno"].ToString() + " \n\nDated:________________") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p20);
                Paragraph p21 = new Paragraph(str: "THE REGIONAL OFFICER\nCENTRAL BOARD OF SECONDARY EDUCATION,\nREGIONAL OFFICE\nAMBIKA COMPLEX, BEHIND SBI COLONY NEAR BRAHMSTHAN\nSHEIKHPURA, BAILEY ROAD PATNA, BIHAR - 800014\n\nSir,\n       With reference to your Confidential Letter No.:RO(PTN)/CONF./H.E./ " + dr["hesub"].ToString() + " /" + dr["heno"].ToString() + " / 2021-MAIN / Dtd: " + date_str + " I hereby accept to act as H.E.") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p21);
                Paragraph p22 = new Paragraph(str: "      My appointment and any information which may come to my notice  during  the discharge of my duties as Head Examiner will be kept confidential.  I undertake to do this work with perfect efficiency / accuracy and according to the instructions issued by the Board to HEs from time to time.\n      I CERTIFY THAT I HAVE NO NEAR RELATION INTENDING / APPEARING IN THE SUBJECT(S) AT THE AFORESAID EXAMINATION.  I ALSO CERTIFY THAT I HAVE NOT WRITTEN ANY HELP BOOK OR NOTES FOR THE EXAMINATION OF THE BOARD.  I UNDERTAKE TO COMPLETE THE WORK ENTRUSTED TO ME WITHIN THE STIPULATED TIME / SCHEDULE FIXED BY THE BOARD.\n        I also certify that I have not been appointed as Head Examiner in  any other subject in Class XII or Class X.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p22);
                PdfPTable tbl3 = new PdfPTable(2);
                tbl3.WidthPercentage = 100f;
                tbl3.HorizontalAlignment = 1;
                tbl3.DefaultCell.Border = 0;
                tbl3.AddCell(new Phrase("(Mobile No.): " + dr["hemobile"].ToString() + "\n\nACCOUNT DETAILS AS PER DATA:\nBank Account No.: " + dr["accountno"].ToString() + "\nIFSC Code           : " + dr["ifsccode"].ToString() + ""));
                tbl3.AddCell(new Phrase("Name: " + dr["hename"].ToString() + "\nDesignation: " + dr["he_desig"].ToString() + "\nSchool: " + dr["headd1"].ToString() + "\n             " + dr["headd2"].ToString() + "\n             " + dr["headd3"].ToString() + "\n             " + dr["headd4"].ToString() + "\n             (" + dr["headd5"].ToString() + " - " + dr["hepin"].ToString() + ")"));
                pdoc.Add(tbl3);
                Paragraph p24 = new Paragraph(str: "I undertake that information  mentioned above are true and there is no change in above records.") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p24);
                Paragraph p23 = new Paragraph(str: "Yours faithfully,     \n\nSignature ...................................\nName of HE: " + dr["hename"].ToString() + "\nHE No.: " + dr["heno"].ToString() + "                       ") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p23);
                //
                pdoc.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
