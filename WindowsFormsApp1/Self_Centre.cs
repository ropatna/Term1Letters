﻿using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class Self_Centre : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=CBSEPAT\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public Self_Centre()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            string database = Microsoft.VisualBasic.Interaction.InputBox("ENTER NAME OF DATABASE FROM WHICH LETTER HAS TO BE GENERATED", "INPUT DATABASE NAME", "CSSELF2021");
            if (String.IsNullOrEmpty(textBox6.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where cen_no='" + textBox6.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Document doc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(doc, new FileStream("F:\\pdf\\SELF\\" + dr["CSSCHNO"].ToString() + "_" + dr["CEN_NO"].ToString() + "_CS.pdf", FileMode.Create));
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
                rosign.SetAbsolutePosition(470, 160);
                doc.Open();
                doc.Add(header); //Adding Header
                doc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                doc.AddTitle("SELF CENTRE CS Letter");
                void self(Document pdoc)
                {
                    //
                    Paragraph p = new Paragraph("===============================================================================\n");
                    pdoc.Add(p);
                    Paragraph p2 = new Paragraph("MOST IMMEDIATE / SPPED POST", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                    pdoc.Add(p2);
                    Paragraph p3 = new Paragraph("**************************") { Alignment = Element.ALIGN_CENTER };
                    pdoc.Add(p3);
                    Paragraph p4 = new Paragraph(str: "Ref.No.:CBSE/RO(PTN)/CONF./CS(Self)/ " + dr["CEN_NO"].ToString() + " /2021/                                                   Date:" + date_str) { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p4);
                    Paragraph p5 = new Paragraph(str: dr["CSSCHNO"].ToString() + "\n" + dr["CSNAME"].ToString() + "\n Vice Principle / Sr. PGT of \n" + dr["ADD1_"].ToString() + "\n" + dr["ADD2_"].ToString() + "\n" + dr["ADD3_"].ToString() + "\n" + dr["ADD4_"].ToString() + "\n" + dr["ADD5_"].ToString() + "\nPIN: " + dr["PIN_"].ToString() + "\n\n") { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p5);
                    Paragraph p6 = new Paragraph(str: "Sub:  Intimation  regarding  appointment  of  Centre  Superintendent  for (SELF CENTRE) AISSE(X) / AISSCE(XII) Main 2021 at Centre No." + dr["cen_no"].ToString() + "\n\n") { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p6);
                    Paragraph p7 = new Paragraph(str: "Dear Sir/Madam,\n\n") { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p7);
                    Paragraph p8 = new Paragraph(str: "      The AISSE(X)/AISSCE(XII)Main 2021 will commence from  4TH MAY 2021 in the  Morning  session  at  10.30 A.M.  as per  respective  Date  Sheet of Examinations.Keeping in view of your teaching experience  and  data submitted by  the  schools  for  Exam 2021, I  am  to  inform  you that  you  have  been appointed as Centre Superintendent  for  Board's  AISSE(X)/AISSCE(XII)  2021 at the following Examination Centre:-\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p8);
                    PdfPTable tbl1 = new PdfPTable(2);
                    tbl1.HorizontalAlignment = 1;
                    tbl1.DefaultCell.Border = 0;
                    tbl1.AddCell(new Phrase(dr["cen_no"].ToString() + "\nTHE PRINCIPAL\n" + dr["add1"].ToString() + "\n" + dr["add2"].ToString() + "\n" + dr["add3"].ToString() + "\n" + dr["add4"].ToString() + "\n" + dr["add5"].ToString() + "\nPIN: " + dr["pin"].ToString() + "\n"));
                    tbl1.AddCell(new Phrase(dr["MOBILE"].ToString() + "\n" + dr["CENTemail"].ToString() + "\n"));
                    pdoc.Add(tbl1);
                    Paragraph p9 = new Paragraph(str: "      The Question Papers of your Centre will be stored  at the following Bank which may please be kept confidential:\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p9);
                    PdfPTable tbl2 = new PdfPTable(2);
                    tbl2.HorizontalAlignment = 1;
                    tbl2.DefaultCell.Border = 0;
                    tbl2.AddCell(new Phrase(dr["bnk_add1"].ToString() + "\n" + dr["bnk_add2"].ToString() + "\n" + dr["bnk_add3"].ToString() + "\n" + dr["bnk_city"].ToString() + "\n" + dr["bnk_dist"].ToString() + "\nPIN: " + dr["bnk_pin"].ToString() + "\n"));
                    tbl2.AddCell(new Phrase("Ph.(O): " + dr["o_ph1"].ToString() + "\nPh.(R): " + dr["r_ph1"].ToString() + "\n"));
                    pdoc.Add(tbl2);
                    Paragraph p10 = new Paragraph(str: "      The  bank  will  deliver  the Question Papers of the subject to  the Centre Supdt. or  his / her  representative on the day(s) of  Examination at an appropriate time  so that the same  would  reach the Exam Centre at 9:30 A.M. positively.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p10);
                    pdoc.NewPage();
                    pdoc.Add(header); //Adding Header
                    pdoc.Add(footer); //Adding Foter
                    Paragraph p11 = new Paragraph(str: "      The number of candidates of Class X/XII allotted at your Centre is as under:\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p11);
                    PdfPTable tbl3 = new PdfPTable(2);
                    tbl3.HorizontalAlignment = 2;
                    tbl3.DefaultCell.Border = 0;
                    tbl3.AddCell(new Phrase("AISSE  (CLASS-X): "));
                    tbl3.AddCell(new Phrase(dr["tot10"].ToString()));
                    tbl3.AddCell(new Phrase("AISSCE  (CLASS - XII): "));
                    tbl3.AddCell(new Phrase(dr["tot12"].ToString()));
                    pdoc.Add(tbl3);
                    Paragraph p12 = new Paragraph(str: "      You are requested to make necessary arrangements at the Examination Centre with prior consultation with the Principal of the school.  This office has also requested the Principal of the School to extend full cooperation to you for smooth & fair conduct of examination.\n\n      Both the  All  India  Senior  School  Certificate  Main  Examination (Class XII) and Secondary School  Examination  Main(Class X) will be held at the Exam Centre and only one Centre Supdt. will be appointed for the purpose.  However,  in  case  the  number  of candidates allotted is more than 250, one Deputy Centre  Superintendent can be  appointed as per  rules given  in the guidelines  to  the  Centre  Superintendent  to  be supplied to you alongwith Centre material.  A remuneration of Rs. 250 / -per day  and  conveyance charges @Rs. 250 / -per day will be paid to the Deputy Superintendent.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p12);
                    Paragraph p13 = new Paragraph(str: "      Remuneration payable to the Centre Superintendent (outside) is @ Rs. 350 / -per day.  In  case  of  local  Supedrintendent conveyance charges  @ Rs. 250 / -per day is also payable.  They will not be entitled to claim TA / DA though they  might  be eligible  for it under the Centre Government / Board rules due to their residence being in the area of different Municipal Board other than that  of  the  Examination  Centre.  Suburbs will be treated within the local limits  for  this   purpose.  No TA / DA will be payable to local Centre Supdts.  Other staff at the Exam Centre may be deployed as given under:") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p13);
                    Paragraph p14 = new Paragraph(str: "       REMUNERATION\\CONVEYANCE IS ADMISSIBLE AS PER LAST YEAR OR AS PER CS GUIDELINES-2021:\n        =================================================================================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 10)) { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p14);
                    PdfPTable tbl4 = new PdfPTable(2);
                    tbl4.HorizontalAlignment = 0;
                    tbl4.DefaultCell.Border = 0;
                    tbl4.AddCell(new Phrase("a) In a hall or big rooms"));
                    tbl4.AddCell(new Phrase(":  - One Asstt. Supdt. for every 12\n   Candidates or part thereof."));
                    tbl4.AddCell(new Phrase("b) In smaller rooms having upto\n   40 candidates"));
                    tbl4.AddCell(new Phrase(":  - Two Asstt. Supdts. in each room.\n   If the no. of candidate exceeds 20."));
                    tbl4.AddCell(new Phrase("c) Clerk"));
                    tbl4.AddCell(new Phrase(":  - One for each centre"));
                    tbl4.AddCell(new Phrase("CLASS IV EMPLOYEES"));
                    tbl4.AddCell(new Phrase(""));
                    tbl4.AddCell(new Phrase("----------------------"));
                    tbl4.AddCell(new Phrase(""));
                    tbl4.AddCell(new Phrase("  Upto 20 candidates"));
                    tbl4.AddCell(new Phrase(":  - One    |Remuneration @ Rs. 100/-"));
                    tbl4.AddCell(new Phrase("  Between 20 to 100 candidates"));
                    tbl4.AddCell(new Phrase(":  - Two    |per day is payable."));
                    tbl4.AddCell(new Phrase("  Between 101 to 400 candidates"));
                    tbl4.AddCell(new Phrase(":  - Three  |"));
                    tbl4.AddCell(new Phrase("  Above 401 candidate"));
                    tbl4.AddCell(new Phrase(":  - Four   |"));
                    pdoc.Add(tbl4);
                    Paragraph p15 = new Paragraph(str: "      Remuneration to Asstt. Supdt. @ Rs. 200/- and conveyance to outsider @ Rs. 150 / -per day.  Asstt.Supdt.(Invigiliator)  for physically challanged candidates  getting 60 mins.extra  time will not be paid extra remuneration.  The clerk will be paid remuneration @Rs. 200 / -per day.\n\n      As  far  as possible, clerk  and  class  IV  employees be appointed from the school itself.  Persons from outside only  be  appointed  in  case the  Principal  of  the  school is not  in a position  to spare  the clerk and the  required  number  of  class IV employees.  In no  case, the  Centre Superintendent should take his/her own Asstt.Supdt.or other staff with him/her if they become entitled for TA/DA.") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p15);
                    pdoc.NewPage();
                    pdoc.Add(header); //Adding Header
                    pdoc.Add(footer); //Adding Foter
                    Paragraph p16 = new Paragraph(str: "      Asstt.  Superintendents should  be  appointed from  teachers of the school(PGT / TGT / PRT)  and be reliable.  But teachers of a particular  subject should  not  be  appointed as Assistant Superintendent  on  the  day  of examination   of   the   said   subject.\n\n      YOU WILL ALSO TAKE AN  UNDERTAKING FROM THE  PERSONS  SO  APPOINTED  THAT NO NEAR  RELATIVE OF  HIS / HER IS APPEARING AT THE EXAMINATION  CONCERNED.  THEY SHOULD  NOT  BELONG  TO  THE  SCHOOL  FROM  WHICH  THE  STUDENTS  ARE TAKING EXAMINATION AT THE CENTRE.  Persons living at far  off  places should  not be appointed as Asstt.Supdt. and  no TA / DA is admissible to Asstt.Supdt.\n\n      Day  to  day details of deployment of  invigilators  be  maintained and preferably duties may be rotated.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p16);
                    Paragraph p17 = new Paragraph(str: "      A sum of Rs. 10/- for infrastructure usage charges  to  examination centre including sitting arrangement will be payable for maximum number  of candidates registered during the examination on any  day.  In addition to it Rs. 15 / -per candidate for maximum number of candidates registered during the examination on any day, on account of stationery, packing materials etc. shall be admissible.  However it does not include conveyance charges  for depositing / despatch of  Answer books and Postage  charges  of  the parcel. Rs 3 / -per candidate alloted to the center for whole exam towards printing centre material, attendance sheet etc.shall be admissible.\n\n      To  witness the opening of Question  Papers  and  packing  of Answer Sheets   one   teacher  from each school  of  Examinees  at   your  Centre may be  called  on  the days of Examination.  The same teacher may be appointed as invigiliator, however such teachers should not be deployed on Invigilation duty  in the  hall / room  where  their   own  school  candidates  are  taking examination.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p17);
                    Paragraph p18 = new Paragraph(str: "      THE CENTRE SUPERINTENDENT IS  PERSONALLY  REQUIRED  TO  ENSURE THAT, AFTER THE EXAMINATION IS OVER, ANSWER  SCRIPTS AND  RELATED MATERIALS OF  THE DAY SHOULD BE SEALED IN GOOD QUALITY CLOTH, AND STRICTLY BE DESPATCHED ON THE SAME DAY TO THE REGIONAL OFFICER, CBSE PATNA BY  INSURED  SPEED POST  ONLY IN SEPERATE SEALED COVER.  IN NO CASE OR CIRCUMSTANCES IT IS  TO  BE  KEPT AT THE CENTRE SCHOOL OVERNNIGHT AND ALSO NOT TO  BE  DESPATCHED  THROUGH  COMMERCIAL PARCEL / RAILWAY PARCEL SERVICE OR PRIVATE COURIER.\n      ATTENDANCE SHEET FOR SSE(X)/AISSCE(XII) Main 2021 SHOULD BE SUBMITTED ON  THE LAST  DAY OF THE EXAM BY  HAND DEPUTING  ONE TEACHER  OF  THE CENTRE.  THE CENTRE SUPERINTENDENT SHALL PERSONALLY BE HELD RESPONSIBLE  IN  CASE  THE ATTENDANCE  SHEETS OF THE CANDIDATES  APPEARED  FROM  THEIR  CENTRE  WILL  BE RECEIVED LATE IN THE REGIONAL OFFICE.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p18);
                    pdoc.NewPage();
                    pdoc.Add(header); //Adding Header
                    pdoc.Add(footer); //Adding Foter
                    Paragraph p19 = new Paragraph(str: "       IMPORTANT INSTRUCTIONS / POINTS: FOR DESPATCH OF SEALED ANSWER BOOK PARCELS\n        =================================================================================", FontFactory.GetFont(FontFactory.TIMES_BOLD, 10)) { Alignment = Element.ALIGN_LEFT };
                    pdoc.Add(p19);
                    PdfPTable tbl5 = new PdfPTable(2);
                    tbl5.WidthPercentage = 100f;
                    tbl5.SetWidths(new float[] { 10f, 130f });
                    tbl5.HorizontalAlignment = 0;
                    tbl5.DefaultCell.Border = 0;
                    tbl5.AddCell(new Phrase("(1)."));
                    tbl5.AddCell(new Phrase("THE CENTRE SUPDT. IS STRICTLY INSTRUCTED THAT HE SHOULD WRITE THE ADDRESS ON  THE SEALED ANSWER BOOK PARCELS OF CLASS X WITH RED  COLOUR AND ON THE PARCELS OF CLASS XII WITH BLUE COLOUR ALONGWITH THE DETAILS VIZ. DATE  OF EXAM, CLASS AND SUBJECT SO THAT IT IS EASILY DISTINGUISHABLE  WHETHER THE PARCELS CONTAIN THE ANSWER BOOKS OF CLASS X OR CLASS XII."));
                    tbl5.AddCell(new Phrase("(2)."));
                    tbl5.AddCell(new Phrase("ON  ALL   OTHER   PARCELS   CONTAINING   MATERIAL   NOT  RELATED  TO  CSO (SECRECY WORK), THE CENTRES SHALL WRITE THE ADDRESS IN BLUE COLOUR IN BOLD LETTERS BUT IN THE  BOTTOM, HE CAN WRITE  WITHIN  BRACKET  'NOT  FOR CSO' SO THAT THIS MATERIAL  COULD  IMMEDIATELY BE SEGREGATED WHEN RECEIVED  IN THE REGIONAL OFFICE AND IS NOT BE HANDED OVER TO CSOs(SECRECY TEAM)."));
                    tbl5.AddCell(new Phrase("(3)."));
                    tbl5.AddCell(new Phrase("SUBJECT CODE, NAME OF SUBJECT AND DATE OF EXAMINATION MUST BE MENTIONED ON THE ANSWER BOOK PARCELS CLEARLY AND THE ANSWER BOOKS OF EACH SUBJECT SHOULD BE PACKED SEPERATELY.  BESIDES THIS, THE NAMES OF THE SIGNATORY  THOSE  WHO WILL BE PRESENT AT THE TIME OF PACKING OF ANSWER BOOK  PARCELS  SHOULD  BE DISCLOSED CLEARLY."));
                    tbl5.AddCell(new Phrase("(4)."));
                    tbl5.AddCell(new Phrase("THE CENTRE SUPDT. MUST ENSURE/ASCERTAIN CORRECTNESS OF FILLED-IN ATTENDANCE SHEET VIZ SERIAL NUMBER OF ANSWER BOOKS, QUESTION PAPER CODE & SET NUMBER ALONGWITH SIGNATURE OF THE CANDIDATE."));
                    pdoc.Add(tbl5);
                    Paragraph p20 = new Paragraph(str: "      Answer   Books,  Supplementary   Answer  Books   and   other   related Material have already been sent  to  Centre School, which may kindly  be  kept under your safe custody.  After the examinations are over, balance Answer Books (Main & Supplementary) etc. may  please be  handed over  to  the  Principal of Centre  School  with  proper account for safe custody.  FOR  STRICT  COMPLIANCE   AND  ADMISIBILITY  OF   RE-IMBURSEMENT  OF EXPENDITURE, REMUNERATIONS AND  HONOURARIUM, YOU  ARE  REQUESTED  TO PERUSE GUIDELINES FOR CENTRE SUPERINTENDENT.\n      The School/Centre must ensure social distancing of candidate at their centre in light  of guidelines  prescribed  by Govt.of India  for  COVID - 19 pandemic.  Guideline  to prevent  COVID - 19  pandemic  during  examination  of AISSE / AISSCE - 2021 shall be made available to the centres in CS Guidelines - 2021.") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p20);
                    Paragraph p21 = new Paragraph(str: "      Your acceptance as Centre Supdt. may be sent to  the  undersigned  in the enclosed  performa  by  return Speed Post / email on ropatna.cbse@nic.in, abcell.cbseropatna@gmail.com latest by 23 / 03 / 2021 positively.  I hope that you will take due care in  making necessary arrangements and extend your  full  cooperation  to  the  Regional  office  of  the  Board for  smooth & fair conduct of Main Examinations 2021, so that we  may fulfill our commitments in the public interest.") { Alignment = Element.ALIGN_JUSTIFIED };
                    pdoc.Add(p21);
                    Paragraph p22 = new Paragraph(str: "      Yours faithfully,\n") { Alignment = Element.ALIGN_RIGHT };
                    pdoc.Add(p22);
                }
                self(doc);
                doc.Add(rosign);
                Paragraph p23 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                doc.Add(p23);
                doc.NewPage();
                doc.Add(header2); //Adding Header
                doc.Add(footer); //Adding Foter
                Paragraph p25 = new Paragraph(str: "ACCEPTANCE OF CENTRE SUPERINTENDENT OF SELF CENTRES 2021\n=================================================================\nR E G I S T E R E D / E M A I L", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                doc.Add(p25);
                Paragraph p26 = new Paragraph(str: "CENTRE NO.  " + dr["cen_no"].ToString() + "       \n") { Alignment = Element.ALIGN_RIGHT };
                doc.Add(p26);
                PdfPTable tbl6 = new PdfPTable(2);
                tbl6.SetWidths(new float[] { 100f, 80f });
                tbl6.HorizontalAlignment = 0;
                tbl6.DefaultCell.Border = 0;
                tbl6.DefaultCell.MinimumHeight = 20f;
                tbl6.AddCell(new Phrase(str: "The Regional Officer,"));
                tbl6.AddCell(new Phrase(str: ""));
                tbl6.AddCell(new Phrase(str: "Central Board of Secondary Education,"));
                tbl6.AddCell(new Phrase(str: "Email: ropatna.cbse@nic.in"));
                tbl6.AddCell(new Phrase(str: "Regional Office"));
                tbl6.AddCell(new Phrase(str: "abcell.cbseropatna@gmail.com"));
                tbl6.AddCell(new Phrase(str: "Ambika Complex,Behind State Bank Colony"));
                tbl6.AddCell(new Phrase(str: ""));
                tbl6.AddCell(new Phrase(str: "Near Brahmsthan,Sheikhpura, Bailey Road"));
                tbl6.AddCell(new Phrase(str: ""));
                tbl6.AddCell(new Phrase(str: "Patna (Bihar) - 800 014"));
                doc.Add(tbl6);
                Paragraph p27 = new Paragraph(str: "\nSir,\n\n       With reference to your letter No. CBSE/RO(PTN)/CONF./CS(Self)/" + dr["cen_no"].ToString() + " / dated: " + date_str + ". I  hereby  express  my  willingness to act  as Centre  Supdt.  for  Centre  No. " + dr["cenno"].ToString() + ".I  shall  conduct  the AISSE/AISSCE Main 2021 as  per  the instructions/guidelines  issued  by the Board.\n\n       I  hereby  certify  that  none of my   near relative is appearing  in  the aforesaid Examinations of the Board.") { Alignment = Element.ALIGN_JUSTIFIED };
                doc.Add(p27);
                Paragraph p24 = new Paragraph(str: "Yours faithfully,     \n\nSignature .................................................\n") { Alignment = Element.ALIGN_RIGHT };
                doc.Add(p24);
                PdfPTable tbl7 = new PdfPTable(3);
                tbl7.SetWidths(new float[] { 100f, 50f, 100f });
                tbl7.WidthPercentage = 100f;
                tbl7.HorizontalAlignment = Right;
                tbl7.DefaultCell.Border = 0;
                tbl7.AddCell(new Phrase(str: "Name and Address of School:\n\n" + dr["ADD1_"].ToString() + "\n" + dr["ADD2_"].ToString() + "\n" + dr["ADD3_"].ToString() + "\n" + dr["ADD4_"].ToString() + "\n" + dr["ADD5_"].ToString() + "\nPIN: " + dr["PIN_"].ToString() + "\nEmail  Id.:  " + dr["cs_email"].ToString() + "\n\n"));
                tbl7.AddCell(new Phrase(str: "\n\n"));
                tbl7.AddCell(new Phrase(str: "\nName: " + dr["CSNAME"].ToString() + "\nDesig.:  VICE PRINCIPAL / SR. PGT \n\nMobile No. for using\nCMTM-App :" + dr["cs_mobile"].ToString() + "\n\n"));
                doc.Add(tbl7);
                doc.Close();
                //
                //
                Document doc2 = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter2 = PdfWriter.GetInstance(doc2, new FileStream("F:\\pdf\\SELF\\" + dr["CSSCHNO"].ToString() + "_" + dr["CEN_NO"].ToString() + "_school.pdf", FileMode.Create));
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(90f, 30f);
                rosign.SetAbsolutePosition(470, 440);
                doc2.Open();
                doc2.Add(header); //Adding Header
                doc2.Add(footer); //Adding Foter
                doc2.AddTitle("SELF CENTRE CS Letter");
                self(doc2);
                Paragraph p31 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                doc2.Add(p31);
                doc2.NewPage();
                doc2.Add(header); //Adding Header
                doc2.Add(footer); //Adding Foter
                Paragraph p28 = new Paragraph(str: "\n\n\nCopy To:\n\nThe Principal,(" + dr["CSSCHNO"].ToString() + ")\n" + dr["ADD1_"].ToString() + "\n" + dr["ADD2_"].ToString() + "\n" + dr["ADD3_"].ToString() + "\n" + dr["ADD4_"].ToString() + "\n" + dr["ADD5_"].ToString() + "\nPIN: " + dr["pin_"].ToString() + "\n") { Alignment = Element.ALIGN_LEFT };
                doc2.Add(p28);
                Paragraph p29 = new Paragraph(str: "       With  the  request  to  relieve  above  Vice  Principal / PGT  from the School to act as Centre Superintendent at the above fixed Centre for 2021 Main Examinations as per  the  undertaking / data forwarded  by you.  The status of releiving of the teacher concern must be confirmed to the  undersigned  on priority.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                doc2.Add(p29);
                doc2.Add(rosign);
                Paragraph p30 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                doc2.Add(p30);
                doc2.Close();
                //
                //
                Document doc3 = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter3 = PdfWriter.GetInstance(doc3, new FileStream("F:\\pdf\\SELF\\" + dr["centschno"].ToString() + "_centre.pdf", FileMode.Create));
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(90f, 30f);
                rosign.SetAbsolutePosition(470, 440);
                doc3.Open();
                doc3.Add(header); //Adding Header
                doc3.Add(footer); //Adding Foter
                doc3.AddTitle("SELF CENTRE CS Letter");
                self(doc3);
                Paragraph p32 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                doc3.Add(p32);
                doc3.NewPage();
                doc3.Add(header); //Adding Header
                doc3.Add(footer); //Adding Foter
                Paragraph p33 = new Paragraph(str: "\n\n\nCopy To:\n\nThe Principal,(" + dr["centschno"].ToString() + ")\n" + dr["add1"].ToString() + "\n" + dr["ADD2"].ToString() + "\n" + dr["ADD3"].ToString() + "\n" + dr["ADD4"].ToString() + "\n" + dr["ADD5"].ToString() + "\nPIN: " + dr["pin"].ToString() + "\n") { Alignment = Element.ALIGN_LEFT };
                doc3.Add(p33);
                Paragraph p34 = new Paragraph(str: "       For information and with the request to extend full co-operation  to Centre Supdt. at  your  school  centre  for  smooth and   fair  conduct  of examination 2021.  It may also be noted that any deviation in regard with conduct of examination  may  leads  the  administrative action against the defaulting.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                doc3.Add(p34);
                doc3.Add(rosign);
                Paragraph p35 = new Paragraph(str: "\n\n(JAGADISH BARMAN) \nREGIONAL OFFICER\n") { Alignment = Element.ALIGN_RIGHT };
                doc3.Add(p35);
                doc3.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
