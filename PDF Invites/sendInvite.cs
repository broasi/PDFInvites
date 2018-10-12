using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using Novacode;
using System.Text.RegularExpressions;

namespace PDF_Invites
{
    class sendInvite
    {
        //DataClasses1DataContext dc = new DataClasses1DataContext();

        static void Main(string[] args)
        {            
            var apex_JobNo = "";
            var NFP_code = "";

            //test.MainTest();

            if (!Parameters.staticRun)
            {
                using (SqlConnection connection = new SqlConnection("Data Source=SQLCLUSTER\\SQLCLUSTER;Initial Catalog=Stratadial;Integrated Security=SSPI"))
                {
                    string SQLString = "SELECT DISTINCT [Apex Job Number] as jobno, [NFP Code/Janssen ID] as nfpCode FROM JB_INVITE_TRACKER WHERE inviteSent = 'False'";
                    SqlCommand comm = new SqlCommand(SQLString, connection);
                    comm.Connection.Open();
                    SqlDataReader r =
                    comm.ExecuteReader(CommandBehavior.CloseConnection);
                    while (r.Read())
                    {
                        apex_JobNo = r["jobno"].ToString().Trim();
                        NFP_code = r["nfpCode"].ToString().Trim();
                        try
                        {
                            create_BodyText(apex_JobNo, NFP_code);
                        }
                        catch { }
                    }
                    r.Close();
                }
            }
            else
            {
                apex_JobNo = Parameters.JobNo;
                NFP_code = Parameters.NFPId;
                create_BodyText(apex_JobNo, NFP_code);
            }
            
        }

        protected static void create_BodyText(string apex_JobNo, string NFP_code)
        {
            var query = "WITH S AS " +
                        "(" +
                         "SELECT topic, SHW_ACJANS28_MEETINGPLANNER.email as mplanner_email, t1.[Meeting Coordinator]," +
                         "isnull([apex job number],'Job Number') as [Apex Job Number], " +
                         "isnull([confirmedspeaker], 'NO CONFIRMED SPEAKER') as [Confirmed Speaker]," +
                         " isnull(t2.[credential 1], '') as [Credential 1], " +
                         " isnull(t2.[credential 2], '') as [Credential 2], " +
                         " isnull(t2.[credential 3], '') as [Credential 3], " +
                         " isnull([dual speaker], '') as [Dual Speaker], " +
                         " isnull(datename(weekday, convert(datetime, [confirmed event date]))+', '+ " +
                         " datename(month, convert(datetime, [confirmed event date]))+' '+ " +
                         " datename(day, convert(datetime, [confirmed event date]))+', '+ " +
                         " datename(Year, convert(datetime, [confirmed event date]))+' ','No Confirmed Date') as [CED], " +
                         //" isnull([time], 'arrival time')+' '+isnull([Zone], '') as [Arrival Time], " +
                         " isnull([time], 'arrival time') as [Arrival Time], " +
                         " isnull([Presentation time], 'No Presentation Time') as [Presentation Time], " +
                         " isnull([zone], 'zone') as [Zone], " +
                         " isnull([Meal and Presentation type], 'No Meal and Presentation type') as [MAPTYPE], " +
                         " isnull([Venuename], 'No Venue') as Venuename, " +
                         " isnull([Venueaddress1], 'No Venue Addr') as Venueaddress1, " +
                         " isnull([Venuecity], 'No Venue City') as [VenueCity], " +
                         " isnull([Venuestate], 'No Venue State') as [VenueState], " +
                         " isnull([Venuezip], 'No Venue Zip') as [VenueZip], " +
                         " isnull([Venuephone], 'No Venue phone') as [Venuephone], " +
                         " isnull([nfp code/janssen id], 'No NFP') as NFP, " +
                         " isnull(convert(varchar(5), month(convert(datetime, [confirmed event date], 101)-10),1)+'/'+convert(varchar(5), day(convert(datetime, [confirmed event date], 101)-10),1)+'/'+convert(varchar(5), year(convert(datetime, [confirmed event date], 101)-10),1), 'NO RSVP DATE') as rsvp " +
                         " from SHW_ACJANS28 as t1 left join shw_acjans28_speaker as t2 on t1.confirmedspeaker = t2.[mail attn] " +
                         //" and t1.[apex job number]=t2.job " +
                         " LEFT OUTER JOIN SHW_ACJANS28_MEETINGPLANNER ON t1.[Meeting Coordinator] = SHW_ACJANS28_MEETINGPLANNER.MeetingPlanner " +
                         " where [apex job number]='" + apex_JobNo + "' and [nfp code/janssen id]='" + NFP_code + "'" +
                         ") " +
                         "SELECT isnull([dual speaker], '') as [Dual Speaker], isnull(P.[credential 1], '') as [Dual Speaker Credential 1], isnull(P.[credential 2], '') as [Dual Speaker Credential 2], isnull(P.[credential 3], '') as [Dual Speaker Credential 3], * FROM S LEFT JOIN shw_acjans28_speaker AS P ON S.[Dual Speaker]=P.[mail attn] AND S.[Apex Job Number]=P.[job]";

            //Console.WriteLine(query);
            //Console.ReadLine();

            var regex = "[\x00-\x08\x0B\x0C\x0E-\x1F]";
            using (SqlConnection connection = new SqlConnection("Data Source=SQLCLUSTER\\SQLCLUSTER;Initial Catalog=Stratadial;Integrated Security=SSPI"))
            {
                SqlCommand comm = new SqlCommand(query, connection);
                comm.Connection.Open();
                SqlDataReader r =
                comm.ExecuteReader(CommandBehavior.CloseConnection);
                if (r.Read())
                {
                    Globals.thenfphold = NFP_code;
                    Globals.thejobhold = r["Apex Job Number"].ToString();
                    Globals.thevenuenamehold = Regex.Replace(r["venuename"].ToString(), regex, String.Empty, RegexOptions.Compiled);
                    Globals.thevenueaddrhold = Regex.Replace(r["venueaddress1"].ToString(), regex, String.Empty, RegexOptions.Compiled); 
                    Globals.thevenuecityhold = Regex.Replace(r["venuecity"].ToString(), regex, String.Empty, RegexOptions.Compiled);
                    Globals.thevenuestatehold = Regex.Replace(r["venuestate"].ToString(), regex, String.Empty, RegexOptions.Compiled);
                    Globals.thevenueziphold = Regex.Replace(r["venuezip"].ToString(), regex, String.Empty, RegexOptions.Compiled);
                    Globals.thevenuephonehold = Regex.Replace(r["venuephone"].ToString(), regex, String.Empty, RegexOptions.Compiled);
                    Globals.thecedhold = r["ced"].ToString();
                    Globals.theatimehold = r["arrival time"].ToString();
                    Globals.theptimehold = r["presentation time"].ToString();
                    Globals.thezonehold = r["zone"].ToString();
                    Globals.themaptypehold = r["MAPTYPE"].ToString();
                    Globals.thersvphold = r["rsvp"].ToString();
                    Globals.theconfirmedspeakerhold = r["Confirmed Speaker"].ToString();
                    //Globals.thecredentials1hold =  stringManipulation.WordWrap(r["Credential 1"].ToString(), 35); 
                    //Globals.thecredentials2hold = stringManipulation.WordWrap(r["Credential 2"].ToString(), 35);
                    //Globals.thecredentials3hold = stringManipulation.WordWrap(r["Credential 3"].ToString(), 35);
                    Globals.thecredentials1hold = r["Credential 1"].ToString();
                    Globals.thecredentials2hold = r["Credential 2"].ToString();
                    Globals.thecredentials3hold = r["Credential 3"].ToString();
                    Globals.thedualspeakerhold = r["Dual Speaker"].ToString();
                    Globals.thedualspeakercredentials1hold = r["Dual Speaker Credential 1"].ToString();
                    Globals.thedualspeakercredentials2hold = r["Dual Speaker Credential 2"].ToString();
                    Globals.thedualspeakercredentials3hold = r["Dual Speaker Credential 3"].ToString();
                    //Globals.thedualspeakercredentials1hold = stringManipulation.WordWrap(r["Dual Speaker Credential 1"].ToString(), 35);
                    //Globals.thedualspeakercredentials2hold = stringManipulation.WordWrap(r["Dual Speaker Credential 2"].ToString(), 35);
                    //Globals.thedualspeakercredentials3hold = stringManipulation.WordWrap(r["Dual Speaker Credential 3"].ToString(), 35);
                    Globals.topic = r["topic"].ToString();
                    Globals.mplanner_email = r["mplanner_email"].ToString();
                }
                r.Close();
            }
                        
            // Get all the Description/s of the topic that match
            var path = "\\\\fscluster\\webdata\\UPLOADED_FILES\\Medica\\Invitations\\App\\";

            List<Program> prg = null;
            using (StreamReader r = new StreamReader(path + "Data\\program.json"))
            {
                string json = r.ReadToEnd();
                prg = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Program>>(json);
            }

            foreach (var item in prg)
            {
                foreach (var obj in item.all_topics)
                {
                    //try
                    //{
                        if (obj.topic == Globals.topic.Trim() && item.active == "True")
                        {
                            Globals.topic = obj.topic;
                            Globals.front = item.front;
                            Globals.back = item.back;
                            Globals.orientation = item.orientation;
                            Globals.assembly = item.assembly;
                            Globals.x1 = item.x1;
                            Globals.y1 = item.y1;
                            Globals.x2 = item.x2;
                            Globals.y2 = item.y2;
                            Globals.font_size = item.font_size;
                            Globals.no_of_columns = item.no_of_columns;
                            Globals.description = item.description;
                            Globals.brush_color = item.brush_color;
                            Globals.cred = item.cred;
                            Globals.date = item.date;
                            Globals.venue = item.venue;
                            Globals.present = item.present;
                            Globals.mtgcd = item.mtgcd;
                            Globals.reg = item.reg;
                            Globals.rsvp = item.rsvp;
                            Globals.active = item.active;
                            Globals.page3 = item.page3;
                            Globals.page4 = item.page4;
                            Globals.textWidthLimit = item.textWidthLimit;
                            compose_text(obj.topic, item.description, Globals.thenfphold, Globals.thejobhold, Globals.thevenuenamehold, Globals.thevenueaddrhold, Globals.thevenuecityhold, Globals.thevenuestatehold, Globals.thevenueziphold, Globals.thevenuephonehold, Globals.thecedhold, Globals.theatimehold, Globals.theptimehold, Globals.themaptypehold, Globals.thersvphold, Globals.theconfirmedspeakerhold, Globals.thecredentials1hold, Globals.thecredentials2hold, Globals.thecredentials3hold, Globals.mplanner_email);                        
                    }
                   // }
                   // catch { }
                }
            }
        }

        protected static void compose_text(string topic, string desc, string thenfphold, string thejobhold, string thevenuenamehold, string thevenueaddrhold, string thevenuecityhold, string thevenuestatehold, string thevenueziphold, string thevenuephonehold, string thecedhold, string theatimehold, string theptimehold, string themaptypehold, string thersvphold, string theconfirmedspeakerhold, string thecredentials1hold, string thecredentials2hold, string thecredentials3hold, string mplanner_email)
        {

            //Content here

            var txtbody = "\n";
            var txtbodyspkr = "\n";
            Globals.speakerBlock = "";
            Globals.dateBlock = "";
            Globals.venueBlock = "";
            Globals.meetingCodeBlock = "";
            Globals.registrationBlock = "";
            Globals.rsvpBlock = "";
            Globals.speakerBlock = "\n";

            //PRESENTED BY
            if (Globals.present == "1")
            {
                txtbodyspkr += "PRESENTED BY\n";
                Globals.speakerBlock += "PRESENTED BY\n";
            }

            //Speaker and Creds

            txtbodyspkr += theconfirmedspeakerhold.Trim() + "\n";
            Globals.speakerBlock += theconfirmedspeakerhold.Trim() + "\n";

            if (Globals.cred == "1" || Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
            {
                if (thecredentials1hold == "") { }
                else
                {
                    txtbodyspkr += thecredentials1hold.Trim() + "\n";
                    Globals.speakerBlock += stringManipulation.WordWrap(thecredentials1hold.Trim(), Globals.textWidthLimit); //thecredentials1hold.Trim() + "\n";
                }
            }
            if (Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
            {
                if (thecredentials2hold == "") { }
                else
                {
                    txtbodyspkr += thecredentials2hold.Trim() + "\n";
                    Globals.speakerBlock += stringManipulation.WordWrap(thecredentials2hold.Trim(), Globals.textWidthLimit);  //thecredentials2hold.Trim() + "\n";
                }
            }
            if (Globals.cred == "3" || Globals.cred == "4")
            {
                if (thecredentials3hold == "") { }
                else
                {
                    txtbodyspkr += thecredentials3hold.Trim() + "\n";
                    Globals.speakerBlock += stringManipulation.WordWrap(thecredentials3hold.Trim(), Globals.textWidthLimit); //thecredentials3hold.Trim() + "\n";
                }
            }
            //For Dual Speaker
            if (Globals.thedualspeakerhold.Trim() != "")
            {
                txtbodyspkr += "\n" + Globals.thedualspeakerhold.Trim() + "\n";
                Globals.speakerBlock += "\n" + Globals.thedualspeakerhold.Trim() + "\n";
                if (Globals.cred == "1" || Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thedualspeakercredentials1hold.Trim() == "") { }
                    else
                    {
                        txtbodyspkr += Globals.thedualspeakercredentials1hold.Trim() + "\n";
                        Globals.speakerBlock += stringManipulation.WordWrap(Globals.thedualspeakercredentials1hold.Trim(), Globals.textWidthLimit); //Globals.thedualspeakercredentials1hold.Trim() + "\n";
                    }
                }
                if (Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thedualspeakercredentials2hold.Trim() == "") { }
                    else
                    {
                        txtbodyspkr += Globals.thedualspeakercredentials2hold.Trim() + "\n";
                        Globals.speakerBlock += stringManipulation.WordWrap(Globals.thedualspeakercredentials2hold.Trim(), Globals.textWidthLimit); //Globals.thedualspeakercredentials2hold.Trim() + "\n";
                    }
                }
                if (Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thedualspeakercredentials3hold.Trim() == "") { }
                    else
                    {
                        txtbodyspkr += Globals.thedualspeakercredentials3hold.Trim() + "\n";
                        Globals.speakerBlock += stringManipulation.WordWrap(Globals.thedualspeakercredentials3hold.Trim(), Globals.textWidthLimit); //Globals.thedualspeakercredentials3hold.Trim() + "\n";
                    }
                }
            }

            //txtbodyspkr += "\n";
            txtbodyspkr += "";

            //The Date
            txtbody += thecedhold.Trim() + "\n";
            Globals.dateBlock += "\n" + thecedhold.Trim() + "\n";
            if (Globals.date == "2")
            {
                txtbody += theatimehold.Trim() + " Arrival | " + theptimehold.Trim() + " " + themaptypehold.Trim() + "\n";
                Globals.dateBlock += theatimehold.Trim() + " Arrival | " + theptimehold.Trim() + " " + themaptypehold.Trim() + "\n";
            }
            if (Globals.date == "3" || Globals.date == "4")
            {
                txtbody += theatimehold.Trim() + " Arrival\n" + theptimehold.Trim() + " " + themaptypehold.Trim() + "\n";
                Globals.dateBlock += theatimehold.Trim() + " Arrival\n" + theptimehold.Trim() + " " + themaptypehold.Trim() + "\n";
            }

            txtbody += "\n";

            //The Venue
            txtbody += thevenuenamehold.Trim() + "\n";
            Globals.venueBlock += "\n" + stringManipulation.WordWrap(thevenuenamehold.Trim(), Globals.textWidthLimit) ;
            if (Globals.venue == "2")
            {
                txtbody += thevenueaddrhold.Trim() + " | " + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n";
                Globals.venueBlock += thevenueaddrhold.Trim() + " | " + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n";
            }
            if (Globals.venue == "3")
            {
                txtbody += thevenueaddrhold.Trim() + "\n" + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n";
                Globals.venueBlock += thevenueaddrhold.Trim() + "\n" + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n";
            }
            if (Globals.venue == "4")
            {
                txtbody += thevenueaddrhold.Trim() + "\n" + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n" + thevenuephonehold.Trim() + "\n";
                Globals.venueBlock += thevenueaddrhold.Trim() + "\n" + thevenuecityhold.Trim() + ", " + thevenuestatehold.Trim() + " " + thevenueziphold.Trim() + "\n" + thevenuephonehold.Trim() + "\n";
            }

            //The Meeting Code
            if (Globals.mtgcd == "1")
            {
                txtbody += "\nMeeting Code " + thenfphold.Trim() + "\n";
                Globals.meetingCodeBlock += "\nMeeting Code " + thenfphold.Trim() + "\n";
            }

            //The Registration
            if (Globals.reg == "3")
            {
                txtbody += "\n";
                txtbody += "REGISTRATION\n";
                txtbody += "To register for this program, visit\n";
                txtbody += "www.MyDomeProgramRegistration.com and enter: Meeting Code: " + thenfphold.Trim() + "\n";
                if (Globals.description == "KnowNow-BillingCoding-bifold")
                {
                    Globals.registrationBlock += "\n";
                    Globals.registrationBlock += "\nREGISTRATION\n";
                    Globals.registrationBlock += "Log on to this site: \n mydomeprogramregistration.com \n enter code " + thenfphold.Trim() + ".\n\n";
                    Globals.registrationBlock += "Please note that your e-mail will\n be required for registration.The \ninformation you provide will only be\n used to facilitate your attendance at \nthis program.\n\n";
                    Globals.registrationBlock += "If you have questions about this \nprogram, call 1-877-468-6720. We \nlook forward to your participation \nIn this informative discussion.";
                }
                else
                {
                    Globals.registrationBlock += "\n";
                    Globals.registrationBlock += "REGISTRATION\n";
                    Globals.registrationBlock += "To register for this program, visit\n";
                    Globals.registrationBlock += "www.MyDomeProgramRegistration.com and enter: Meeting Code: " + thenfphold.Trim() + "\n";
                }                
            }
            if (Globals.reg == "4")
            {
                txtbody += "\n";
                txtbody += "\nREGISTRATION\n";
                txtbody += "To register for this program, visit\n";
                txtbody += "www.MyDomeProgramRegistration.com\nMeeting Code: " + thenfphold.Trim() + "\n";
               
                Globals.registrationBlock += "\n";
                Globals.registrationBlock += "\nREGISTRATION\n";
                Globals.registrationBlock += "To register for this program, visit\n";
                Globals.registrationBlock += "www.MyDomeProgramRegistration.com\n and enter: Meeting Code: " + thenfphold.Trim() + "\n";
            }
            //The RSVP
            if (Globals.rsvp == "1")
            {
                txtbody += "\n";
                txtbody += "RSVP's are appreciated by: " + thersvphold.Trim();
                Globals.rsvpBlock += "\n";
                Globals.rsvpBlock += "RSVP's are appreciated by: " + thersvphold.Trim();
            }

            Globals.txtBody = txtbody;
            Globals.txtSpeaker = txtbodyspkr;
            Globals.jobNo = thejobhold;
            Globals.NFP_Code = thenfphold;

            //try
            //{
                generate_Doc(mplanner_email, topic, desc);
            //}
            //catch { }
            generate_PDF(mplanner_email, topic, desc);
        }

        protected static void generate_Doc(string mplanner_email, string topic, string desc)
        {            
            var path = "\\\\fscluster\\webdata\\UPLOADED_FILES\\Medica\\Invitations\\App\\";
                        
            string fileName = path + "Doc\\" + Globals.jobNo + "_" + Globals.description  + "_" + Globals.NFP_Code + ".doc";
           
            var doc = DocX.Create(fileName);

            var textFormatTimesNewRoman = new Novacode.Formatting();
            textFormatTimesNewRoman.FontFamily = new System.Drawing.FontFamily("Times New Roman");
            textFormatTimesNewRoman.Size = 12D;            

            var bodyFormatRegular = new Novacode.Formatting();
            bodyFormatRegular.FontFamily = new System.Drawing.FontFamily("Arial");
            bodyFormatRegular.Size = 12D;

            var bodyFormatBold = new Novacode.Formatting();
            bodyFormatBold.FontFamily = new System.Drawing.FontFamily("Arial");
            bodyFormatBold.Size = 12D;            
            bodyFormatBold.Bold = true;            

            Novacode.Paragraph body = doc.InsertParagraph("Job #: " + Globals.jobNo + "_" + Globals.description + "_" + Globals.thevenuecityhold + ", " + Globals.thevenuestatehold + "_" + Globals.NFP_Code + "_26 copies" + Environment.NewLine + Environment.NewLine, false, textFormatTimesNewRoman);
            body.Alignment = Alignment.left;

            if (Globals.description != "GI Stelara Postcard broadcast" && Globals.description != "National broadcast live" && Globals.description != "National broadcast prerecorded")
            {
                //PRESENTED BY
                if (Globals.present == "1")
                {                
                    body = doc.InsertParagraph("PRESENTED BY", false, bodyFormatBold);
                    body.Alignment = Alignment.center;
                }

                //Speaker and Creds                        
                body = doc.InsertParagraph(Globals.theconfirmedspeakerhold.Trim(), false, bodyFormatBold);
                body.Alignment = Alignment.center;
                if (Globals.cred == "1" || Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thecredentials1hold == "") { }
                    else
                    {
                        body = doc.InsertParagraph(Globals.thecredentials1hold.Trim(), false, bodyFormatRegular);
                        body.Alignment = Alignment.center;
                    }
                }
                if (Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thecredentials2hold == "") { }
                    else
                    {
                        body = doc.InsertParagraph(Globals.thecredentials2hold.Trim(), false, bodyFormatRegular);
                        body.Alignment = Alignment.center;
                    }
                }
                if (Globals.cred == "3" || Globals.cred == "4")
                {
                    if (Globals.thecredentials3hold == "") { }
                    else
                    {
                        body = doc.InsertParagraph(Globals.thecredentials3hold.Trim(), false, bodyFormatRegular);
                        body.Alignment = Alignment.center;
                    }
                }

                //Dual Speaker is not empty
                if (Globals.thedualspeakerhold.Trim() != "")
                {
                    body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);
                    body = doc.InsertParagraph(Globals.thedualspeakerhold.Trim(), false, bodyFormatBold);
                    body.Alignment = Alignment.center;
                    if (Globals.cred == "1" || Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                    {
                        if (Globals.thedualspeakercredentials1hold == "") { }
                        else
                        {
                            body = doc.InsertParagraph(Globals.thedualspeakercredentials1hold.Trim(), false, bodyFormatRegular);
                            body.Alignment = Alignment.center;
                        }
                    }
                    if (Globals.cred == "2" || Globals.cred == "3" || Globals.cred == "4")
                    {
                        if (Globals.thedualspeakercredentials2hold == "") { }
                        else
                        {
                            body = doc.InsertParagraph(Globals.thedualspeakercredentials2hold.Trim(), false, bodyFormatRegular);
                            body.Alignment = Alignment.center;
                        }
                    }
                    if (Globals.cred == "3" || Globals.cred == "4")
                    {
                        if (Globals.thedualspeakercredentials3hold == "") { }
                        else
                        {
                            body = doc.InsertParagraph(Globals.thedualspeakercredentials3hold.Trim(), false, bodyFormatRegular);
                            body.Alignment = Alignment.center;
                        }
                    }
                }
            }

            body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);
                        
            //The Date            
            body = doc.InsertParagraph(Globals.thecedhold.Trim(), false, bodyFormatBold);
            body.Alignment = Alignment.center;
            if (Globals.date == "2")
            {
                if (Globals.description == "National broadcast live" || Globals.description == "National broadcast prerecorded")
                {
                    body = doc.InsertParagraph(Globals.theatimehold.Trim() + " " + Globals.thezonehold + " Arrival | " + Globals.theptimehold.Trim() + " " + Globals.thezonehold + " " + Globals.themaptypehold.Trim(), false, bodyFormatRegular);
                    body.Alignment = Alignment.center;
                }
                else
                {
                    body = doc.InsertParagraph(Globals.theatimehold.Trim() + " Arrival | " + Globals.theptimehold.Trim() + " " + Globals.themaptypehold.Trim(), false, bodyFormatRegular);
                    body.Alignment = Alignment.center;
                }
            }
            if (Globals.date == "3" || Globals.date == "4")
            {
                body = doc.InsertParagraph(Globals.theatimehold.Trim() + " Arrival" + Environment.NewLine + Globals.theptimehold.Trim() + " " + Globals.themaptypehold.Trim(), false, bodyFormatRegular);
                body.Alignment = Alignment.center;                
            }
            body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);

            //The Venue            
            body = doc.InsertParagraph(Globals.thevenuenamehold.Trim(), false, bodyFormatBold);
            body.Alignment = Alignment.center;
            if (Globals.venue == "2")
            {
                body = doc.InsertParagraph(Globals.thevenueaddrhold.Trim() + " | " + Globals.thevenuecityhold.Trim() + ", " + Globals.thevenuestatehold.Trim() + " " + Globals.thevenueziphold.Trim(), false, bodyFormatRegular);
                body.Alignment = Alignment.center;
            }
            if (Globals.venue == "3")
            {
                body = doc.InsertParagraph(Globals.thevenueaddrhold.Trim() + Environment.NewLine + Globals.thevenuecityhold.Trim() + ", " + Globals.thevenuestatehold.Trim() + " " + Globals.thevenueziphold.Trim(), false, bodyFormatRegular);
                body.Alignment = Alignment.center;
            }
            if (Globals.venue == "4")
            {
                body = doc.InsertParagraph(Globals.thevenueaddrhold.Trim() + Environment.NewLine + Globals.thevenuecityhold.Trim() + ", " + Globals.thevenuestatehold.Trim() + " " + Globals.thevenueziphold.Trim() + Environment.NewLine + Globals.thevenuephonehold.Trim(), false, bodyFormatRegular);
                body.Alignment = Alignment.center;
            }
            body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);

            //The Meeting Code
            if (Globals.mtgcd == "1" || Globals.description == "Stelara CD Branded Invite" || Globals.description == "Simponi Aria Nurse invite" || Globals.description == "Rheum Gastro Invite- Stelara unbranded" || Globals.description == "Bio Coordinator invite unbranded" || Globals.description == "Simponi Aria Unbranded invite" || Globals.description == "SIMPONI UC iNVITe" || Globals.description == "Stelara Rheum New Data" || Globals.description == "Nurse Stelara CD Invite")
            {
                body = doc.InsertParagraph("Meeting Code: " + Globals.thenfphold.Trim(), false, bodyFormatRegular);                
                body.Alignment = Alignment.center;
                body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);
            }            

            //The Registration
            if (Globals.reg == "3" || Globals.description == "ImmuneResponse-2sidedPrint")
            {                                
                    body = doc.InsertParagraph("REGISTRATION", false, bodyFormatBold);
                    body.Alignment = Alignment.center;
                    body = doc.InsertParagraph("To register for this program, visit", false, bodyFormatRegular);
                    body.Alignment = Alignment.center;
                    body = doc.InsertParagraph("www.MyDomeProgramRegistration.com and enter: Meeting Code: " + Globals.thenfphold.Trim(), false, bodyFormatRegular);
                    body.Alignment = Alignment.center;
                    body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);
            }
            if (Globals.reg == "4")
            {
                body = doc.InsertParagraph("REGISTRATION", false, bodyFormatBold);
                body.Alignment = Alignment.center;
                body = doc.InsertParagraph("To register for this program, visit", false, bodyFormatRegular);
                body.Alignment = Alignment.center;
                body = doc.InsertParagraph("www.MyDomeProgramRegistration.com " + Environment.NewLine + " and enter: Meeting Code: " + Globals.thenfphold.Trim(), false, bodyFormatRegular);
                body.Alignment = Alignment.center;
                body = doc.InsertParagraph(Environment.NewLine, false, bodyFormatBold);
            }
                        
            //The RSVP
            if (Globals.rsvp == "1" || Globals.description == "Stelara CD Branded Invite" || Globals.description == "Simponi Aria Nurse invite" || Globals.description == "Rheum Gastro Invite- Stelara unbranded" || Globals.description == "Bio Coordinator invite unbranded" || Globals.description == "Simponi Aria Unbranded invite" || Globals.description == "SIMPONI UC iNVITe" ||  Globals.description == "Nurse Stelara CD Invite")
            {                
                body = doc.InsertParagraph("RSVP's are appreciated by: ", false, bodyFormatBold);
                body.Append(Globals.thersvphold.Trim());
                body.Font(new FontFamily("Arial"));
                body.FontSize(12);
                body.Alignment = Alignment.center;
            }
            if (Globals.description == "National broadcast live" || Globals.description == "National broadcast prerecorded")
            {
                body = doc.InsertParagraph("Please RSVP to: ", false, bodyFormatBold);
                body.Append("\nwww.MyDomeProgramRegistration.com");
                body.Font(new FontFamily("Arial"));
                body.FontSize(12);
                body.Alignment = Alignment.center;
            }

            //try
            //{
                doc.Save();
            //}
            //catch { }
            doc.Dispose();            
        }

        protected static void generate_PDF(string mplanner_email, string topic, string desc)
        {          
            var path = "\\\\fscluster\\webdata\\UPLOADED_FILES\\Medica\\Invitations\\App\\";
            //try
             //{
                //List<Program> prg = null;
                //using (StreamReader r = new StreamReader(path + "Data\\program.json"))
                //{
                //    string json = r.ReadToEnd();
                //    prg = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Program>>(json);
                //}

                //var template = prg.First(p => p.description == desc && p.topic == topic);
                //var assembly = template.assembly == "back" ? template.back : template.front;
                var assembly = Globals.assembly == "back" ? Globals.back : Globals.front;

                StringFormat sf = new StringFormat();
                sf.LineAlignment = StringAlignment.Center;
                sf.Alignment = StringAlignment.Center;

                StringFormat sfNear = new StringFormat();
                sfNear.LineAlignment = StringAlignment.Near;
                sfNear.Alignment = StringAlignment.Near;                
                                
                System.Drawing.Text.PrivateFontCollection privateFonts = new System.Drawing.Text.PrivateFontCollection();
                privateFonts.AddFontFile(path + "Fonts\\" + "Karbon-Regular.ttf");
                privateFonts.AddFontFile(path + "Fonts\\" + "Karbon-Bold.ttf");
                privateFonts.AddFontFile(path + "Fonts\\" + "Karbon-RegularItalic.ttf");                
                System.Drawing.Font KarbonFontBold12 = new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Bold);
                System.Drawing.Font KarbonFontBold8 = new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Bold);
                System.Drawing.Font KarbonFontBold6 = new System.Drawing.Font(privateFonts.Families[0], 6, FontStyle.Bold);
                System.Drawing.Font KarbonFontRegular12 = new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular);
                System.Drawing.Font KarbonFontRegular8 = new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular);
                System.Drawing.Font KarbonFontRegular6 = new System.Drawing.Font(privateFonts.Families[0], 6, FontStyle.Regular);
                System.Drawing.Font KarbonFontRegularItalic12 = new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Italic);
                System.Drawing.Font KarbonFontRegularItalic8 = new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Italic);
                System.Drawing.Font KarbonFontRegularItalic6 = new System.Drawing.Font(privateFonts.Families[0], 6, FontStyle.Italic);
                
                //System.Drawing.Font drawFont = new System.Drawing.Font("Arial", Globals.font_size);
                //System.Drawing.Font drawFont2 = new System.Drawing.Font("Arial", 8);
                System.Drawing.Font drawFont = new System.Drawing.Font(privateFonts.Families[0], Globals.font_size, FontStyle.Regular);
                System.Drawing.Font drawFont2 = new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular);

                SolidBrush drawBrush = new SolidBrush(Color.Black);
                if (Globals.brush_color == "white")
                {
                    drawBrush = new SolidBrush(Color.White);
                }

                System.Drawing.Image canvas = Bitmap.FromFile(path + "Template\\" + assembly);
                Graphics gra = Graphics.FromImage(canvas);

                if (Globals.assembly == "front and back")
                {
                    canvas = Bitmap.FromFile(path + "Template\\" + Globals.front);
                    gra = Graphics.FromImage(canvas);

                    System.Drawing.Image canvas2 = Bitmap.FromFile(path + "Template\\" + Globals.back);
                    Graphics gra2 = Graphics.FromImage(canvas2);
                    gra2.DrawString("\nTo register for this program, visit\nwww.MyDomeProgramRegistration.com and enter: Meeting Code: " + Globals.NFP_Code.Trim() + "\n", new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, 1353, 337, sf);
                    canvas2.Save(path + "Assembly\\" + Globals.back, System.Drawing.Imaging.ImageFormat.Png);
                }

                //To get the x,y get the center coordinates of the area to where the text will be placed in the background image
                Point drawPoint1 = new Point(Globals.x1, Globals.y1);
                Point drawPoint2 = new Point(Globals.x2, Globals.y2);

                //replace this if text is too long palit
                Globals.speakerBlock = Parameters.speakerText(Globals.speakerBlock);
                Globals.venueBlock = Parameters.venueText(Globals.venueBlock);
                //Globals.txtSpeaker = Globals.txtSpeaker.Replace("Zeinub Alber, RN, Biologic Coordinator", "Zeinub Alber, RN, \nBiologic Coordinator");
                
                var bodyText = Globals.txtSpeaker + Globals.txtBody;               
                if (Globals.description == "ImmuneResponse-PDF" || Globals.description == "ImmuneResponse-Web")
                {
                    gra.DrawString("\nTo register for this program, visit\nwww.MyDomeProgramRegistration.com \nand enter: Meeting Code: " + Globals.NFP_Code.Trim() + "\n", drawFont, drawBrush, 2284, 1048, sf);
                }
                if (Globals.no_of_columns == 1)
                {
                     gra.DrawString(bodyText, drawFont, drawBrush, drawPoint1, sf);
                }
                if (Globals.no_of_columns == 2)
                {
                    if (Globals.description == "SPA Invite")
                    {
                        gra.DrawString(Globals.speakerBlock +  Globals.dateBlock +  Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.rsvpBlock +  Globals.registrationBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, drawPoint2, sf);
                    }                                    
                    else if (Globals.description == "KnowNow-Biosimilars-web")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thenfphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2934, 617, sf);
                    }
                    else if (Globals.description == "Stelara Traditional" || Globals.description == "Stelara Roundtable")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0],6, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thenfphold, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, 540, 1769, sf);
                        
                    }                    
                    else if (Globals.description == "Biosimilars Invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 18, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thenfphold + '.', new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), new SolidBrush(Color.White), 705, 2590, sf);

                    }
                    else if (Globals.description == "unbranded Biosimilars Invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thenfphold + '.', new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), new SolidBrush(Color.White), 640, 1935, sfNear);

                    }
                    else if (Globals.description == "branded Biosimilars invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thenfphold + '.', new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), new SolidBrush(Color.Black), 493, 1923, sfNear);

                    }
                    else if (Globals.description == "gi Bio Coordinator invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock + Globals.meetingCodeBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, new Point(2100, 1935), sf);
                        gra.DrawString(Globals.rsvpBlock.Replace("RSVP's are appreciated by: ", ""), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 150, 2070, sfNear);
                    }
                    else if (Globals.description == "KnowNow-Biosimilars-bifold" || Globals.description == "KnowNow-Billing-web-connect" || Globals.description == "KnowNow-Billing-bifold print" || Globals.description == "KnowNow-MedicarePayments-bifold" || Globals.description == "KnowNow-MedicarePayments-web" || Globals.description == "KnowNow-MQPP2-bifold" || Globals.description == "Know Now Practice Manager Bifold" || Globals.description == "Know Now Practice Manager web" || Globals.description == "Know now Navigating affordability" || Globals.description == "Know now Navigating affordability-Web")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        if (Globals.description == "Know now Navigating affordability" || Globals.description == "Know now Navigating affordability-Web")
                        {
                            gra.DrawString(Globals.thenfphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2800, 575, sf);
                    }
                        else
                            gra.DrawString(Globals.thenfphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2923, 617, sf);
                    }
                    else if (Globals.description == "National broadcast live" || Globals.description == "National broadcast prerecorded")
                    {
                        gra.DrawString("Program Information", new System.Drawing.Font(privateFonts.Families[0], 16, FontStyle.Bold), drawBrush, 599, 141, sf);
                        gra.DrawString("Date and Time:", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Bold), drawBrush, 599, 217, sf);                        
                        gra.DrawString(Globals.thecedhold + '\n' +Globals.theatimehold + " " + Globals.thezonehold +  " Arrival – " + Globals.theptimehold + " " + Globals.thezonehold +" Dinner & Presentation \n Meeting Code " + Globals.thenfphold, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, 599, 300, sf);                        
                        gra.DrawString("Location:", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Bold), drawBrush, 599, 418, sf);
                        gra.DrawString(Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, 599, 484, sf);
                        gra.DrawString("Please RSVP to:", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Bold), drawBrush, 599, 630, sf);
                        gra.DrawString("www.MyDomeProgramRegistration.com", new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, 599,675, sf);
                    }
                    else
                    {
                        gra.DrawString(Globals.txtSpeaker, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);                        
                        gra.DrawString(Globals.txtBody, drawFont, drawBrush, drawPoint2, sf);
                    }
                }
                if (Globals.no_of_columns == 3)
                {
                    if (Globals.description == "Stelara CD Branded Invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thersvphold, new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 120, 1450, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 1520, 1578, sfNear);
                    }
                    if (Globals.description == "Simponi Aria Nurse invite" ||  Globals.description == "Bio Coordinator invite unbranded" || Globals.description == "Simponi Aria Unbranded invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString("RSVPs ARE APPRECIATED BY\n" + Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 150, 2200, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 1650, 2375, sfNear);
                    }
                    if (Globals.description == "Rheum Gastro Invite- Stelara unbranded")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 155, 2230, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 1650, 2375, sfNear);
                    }
                    if (Globals.description == "Stelara Rheum New Data")
                    {
                        gra.DrawString(Globals.speakerBlock.Replace("PRESENTED BY", "PRESENTED BY\n\n") + "\n\n" + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        //gra.DrawString(Globals.speakerBlock.Replace("PRESENTED BY", "\nPRESENTED BY") + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 516, 2337, sfNear);
                    }
                    if (Globals.description == "SIMPONI UC iNVITe")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 8, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 680, 2178, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 1876, 2330, sfNear);
                    }
                    if (Globals.description == "Stelara CD nurse invite")
                    {
                        gra.DrawString(Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 1886, 456, sf);
                        gra.DrawString(Globals.speakerBlock.Replace("PRESENTED BY",""), new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), new SolidBrush(Color.White), 1406, 1092, sfNear);
                        gra.DrawString(Globals.thersvphold, new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), new SolidBrush(Color.White), 1794, 1878, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), new SolidBrush(Color.White), 1778, 1568, sfNear);
                    }
                    if (Globals.description == "Stelara CD invite" || Globals.description == "Stelara CD invite- Dinner" || Globals.description == "Stelara Biologic Coordinator" || Globals.description == "Stelara CD roundtable")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 581, 1720, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 450, 1895, sfNear);
                    }
                    if (Globals.description == "Stelara CD dinner Invite" || Globals.description == "Stelara CD Case based Considerations Invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 581, 1710, sfNear);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 450, 1885, sfNear);
                    }                    
                    if (Globals.description == "Tremfya invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString("Please RSVP by: " + Globals.thersvphold + "\n", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1278, 2553, sf);
                        gra.DrawString(Globals.NFP_Code.Trim() + ".", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1950, 2650, sfNear);
                    }
                    if (Globals.description == "Branded Simponi Aria invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, drawPoint1, sf);                        
                        gra.DrawString(Globals.NFP_Code.Trim() + ".", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 450, 1915, sfNear);
                    }
                    if (Globals.description == "Simponi PSA and AS invite" || Globals.description == "Branded Simponi AS invite")
                    {                    
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        if (Globals.description == "Branded Simponi AS invite")
                          gra.DrawString(Globals.NFP_Code.Trim() + ".", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 450, 1840, sfNear);
                        else
                          gra.DrawString(Globals.NFP_Code.Trim() + ".", new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 450, 1890, sfNear);
                    }                  
                    if (Globals.description == "Tremfya roundtable invite" || Globals.description == "Tremfya dinner invite" || Globals.description == "Tremfya Branded Bio Coord")
                    {                    
                        gra.DrawString(Globals.thecedhold, new System.Drawing.Font(privateFonts.Families[0], 16, FontStyle.Regular), drawBrush, 1262, 1110, sf);
                        gra.DrawString(Globals.theatimehold, new System.Drawing.Font(privateFonts.Families[0], 16, FontStyle.Regular), drawBrush, 1262, 1222, sf);
                        gra.DrawString(Globals.speakerBlock, new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, 874, 1538, sf);
                        gra.DrawString(Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, 1640, 1545, sf);
                        if (Globals.description == "Tremfya dinner invite")
                        {
                            gra.DrawString(Globals.thersvphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1410, 2609, sfNear);
                            gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1794, 2182, sfNear);
                        }
                        else if (Globals.description == "Tremfya roundtable invite")
                        {
                            gra.DrawString(Globals.thersvphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1778, 2695, sfNear);
                            gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1781, 2365, sfNear);
                        }
                        else
                        {
                            gra.DrawString(Globals.thersvphold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1490, 2632, sf);
                            gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 1771, 2180, sfNear);
                        }
                        
                }
                    if (Globals.description == "Simponi Remicade dual invite")
                    {
                        gra.DrawString(Globals.speakerBlock + Globals.dateBlock + Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, drawPoint1, sf);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 11, FontStyle.Regular), drawBrush, 466, 2015, sfNear);
                    }
                    if (Globals.description == "derm broadcast venue")
                    {
                        gra.DrawString(Globals.dateBlock + "Meeting Code: " + Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, 248, 820, sfNear);
                        gra.DrawString(Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 9, FontStyle.Regular), drawBrush, 248, 1131, sfNear);
                    }
                    if (Globals.description == "gi broadcast")
                    {
                        gra.DrawString(Globals.dateBlock + "Meeting Code: " + Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2697, 655, sfNear);
                        gra.DrawString(Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2691, 980, sfNear);
                    }
                    if (Globals.description == "InfuseU-regional-bifold")
                    {
                        gra.DrawString(Globals.dateBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2042, 325, sf);
                        gra.DrawString(Globals.venueBlock, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2864, 325, sf);
                        gra.DrawString(Globals.theconfirmedspeakerhold + "\n" + stringManipulation.WordWrap(Globals.thecredentials1hold.Trim(), 35)  + stringManipulation.WordWrap(Globals.thecredentials2hold.Trim(), 35) + stringManipulation.WordWrap(Globals.thecredentials3hold.Trim(), 35), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2067, 726, sf);
                        //gra.DrawString(Globals.theconfirmedspeakerhold + "\n" + Globals.thecredentials1hold + "\n" + Globals.thecredentials2hold.Replace("Arthritis & Rheumatology Associates of Palm Beach", "Arthritis & Rheumatology \nAssociates of Palm Beach") + "\n" + Globals.thecredentials3hold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2067, 726, sf);
                        gra.DrawString(Globals.thedualspeakerhold  + "\n" + stringManipulation.WordWrap(Globals.thedualspeakercredentials1hold.Trim(), 35)   + stringManipulation.WordWrap(Globals.thedualspeakercredentials2hold.Trim(), 35) + stringManipulation.WordWrap(Globals.thedualspeakercredentials3hold.Trim(), 35), new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2864, 726, sf);
                        //gra.DrawString(Globals.thedualspeakerhold + "\n" + Globals.thedualspeakercredentials1hold + "\n" + Globals.thedualspeakercredentials2hold.Replace("Arthritis and Rheumatology Assoc of Palm Beach", "Arthritis and Rheumatology \nAssoc of Palm Beach") + "\n" + Globals.thedualspeakercredentials3hold, new System.Drawing.Font(privateFonts.Families[0], 10, FontStyle.Regular), drawBrush, 2864, 726, sf);
                        gra.DrawString(Globals.NFP_Code.Trim(), new System.Drawing.Font(privateFonts.Families[0], 12, FontStyle.Regular), drawBrush, 2938, 995, sfNear);
                    }
                }
                canvas.Save(path + "Assembly\\" + assembly, System.Drawing.Imaging.ImageFormat.Png);
                var filename = Globals.jobNo + "_" + Globals.description + "_" + Globals.NFP_Code;

                var pgSize = new iTextSharp.text.Rectangle(2550, 3300);               //Standard Size Portrait 
                Document doc =   new Document(pgSize, 0f, 0f, 0f, 0f);
                //Document doc = new Document(PageSize.A4.Rotate(), 0f, 0f, 0f, 0f);
                if (Globals.orientation == "LandscapeStandard")
                {
                    pgSize = new iTextSharp.text.Rectangle(3300, 2550);              //Standard Size Landscape
                    doc = new Document(pgSize, 0f, 0f, 0f, 0f);
                }
                if (Globals.orientation == "LandscapeTabloid")                      //Tabloid Landscape
                {                    
                    pgSize = new iTextSharp.text.Rectangle(4913f, 2550f);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);                    
                }
                if (Globals.orientation == "LandscapeTabloidBig")                      //Tabloid Landscape
                {
                    pgSize = new iTextSharp.text.Rectangle(5100f, 3300f);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);
                }
                if (Globals.orientation == "LandscapeTabloidMedium")                      //Tabloid Landscape Medium
                {
                    pgSize = new iTextSharp.text.Rectangle(4800f, 3000f);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);
                }
                if (Globals.orientation == "LandscapeTabloidSmall")                  //Tabloid Landscape Small
                {
                    pgSize = new iTextSharp.text.Rectangle(2700, 1800);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);
                }
                if (Globals.orientation == "PortraitTabloid")                      //Tabloid Portrait
                {
                    pgSize = new iTextSharp.text.Rectangle(1650, 2550);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);
                }
                if (Globals.orientation == "PostCard")                          //PostCard
                {
                    pgSize = new iTextSharp.text.Rectangle(2100, 1500);
                    doc = new iTextSharp.text.Document(pgSize, 0, 0, 0, 0);
                }

                string pdfFilePath = path + "PDF\\";
                string docFilePath = path + "Doc\\";

                PdfWriter writer = PdfAWriter.GetInstance(doc, new FileStream(pdfFilePath +  filename + ".pdf", FileMode.Create));
                //PdfWriter writer = PdfAWriter.GetInstance(doc, ms);
                doc.Open();

                var folder1 = "";
                var folder2 = "";
                string imageURL = "";
                if (Globals.assembly == "front")
                {
                    folder1 = path + "Assembly\\";
                    folder2 = path + "Template\\";
                    imageURL = folder1 + Globals.front;
                }
                else if (Globals.assembly == "back")
                {
                    folder1 = path + "Template\\";
                    folder2 = path + "Assembly\\";
                    imageURL = folder1 + Globals.front;
                }
                else if (Globals.assembly == "front and back")
                {
                    imageURL = path + "Assembly\\" + Globals.front;
                    folder2 = path + "Assembly\\";
                }
                if (Globals.description == "Stelara PSA-dual speaker" || Globals.description == "derm broadcast venue" || Globals.description == "Stelara PSA-single speaker" || Globals.description == "Bio Coordinator invite unbranded" || Globals.description == "053324-160516 Stelara Roundtable PDF-HighRes" || Globals.description == "Stelara Roundtable" || Globals.description == "Stelara Traditional" || Globals.description == "Stelara CD Branded Invite" || Globals.description == "Nurse Stelara CD Invite" || Globals.description == "Stelara CD invite" || Globals.description == "Tremfya invite" ||  Globals.description == "Tremfya Branded Bio Coord" || Globals.description == "Tremfya roundtable invite" || Globals.description == "Tremfya dinner invite"  || Globals.description == "Simponi PSA and AS invite" || Globals.description == "Simponi Remicade dual invite")
                {
                    var note = "note1.png";
                    if (Globals.description == "Tremfya invite" || Globals.description == "derm broadcast venue" || Globals.description == "Tremfya Branded Bio Coord")
                    {
                        note = "note2.png";
                    }
                    if (Globals.description == "Simponi PSA and AS invite")
                    {
                        note = "note3.png";
                    }
                    if (Globals.description == "Nurse Stelara CD Invite" || Globals.description == "Stelara CD invite")
                    {
                        note = "note4.png";
                    }
                    if (Globals.description == "Tremfya roundtable invite" || Globals.description == "Tremfya dinner invite")
                    {
                        note = "note5.png";
                    }
                    if (Globals.description == "Simponi Remicade dual invite")
                    {
                        note = "note6.png";
                    }

                // Note text                                
                iTextSharp.text.Image jpg0 = iTextSharp.text.Image.GetInstance(path + "Template\\" + note);
                    jpg0.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                    jpg0.SpacingBefore = 0f;
                    jpg0.SpacingAfter = 0f;
                    jpg0.Alignment = Element.ALIGN_CENTER;
                    doc.Add(jpg0);
                    //iTextSharp.text.Font timesRoman = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 90, iTextSharp.text.Font.BOLD);
                    //iTextSharp.text.Paragraph pageOneTextOnly = new iTextSharp.text.Paragraph("PLEASE NOTE: IT IS \n REQUIRED THAT YOU \n DELIVER THIS INVITATION WITH A \n COPY OF THE CURRENT \n STELARA® (USTEKINUMAB) \n PRESCRIBING INFORMATION. \n THERE ARE TO BE NO \n EXCEPTIONS TO THIS \n REQUIREMENT.", timesRoman);
                    //pageOneTextOnly.Alignment = Element.ALIGN_LEFT;
                    //pageOneTextOnly.IndentationLeft = 30;                    
                    //pageOneTextOnly.PaddingTop = 0f;
                    //pageOneTextOnly.SpacingBefore = 0f;
                    //pageOneTextOnly.SpacingAfter = 0f;                    
                    //doc.Add(pageOneTextOnly);
                }

                // Page One                                 
                iTextSharp.text.Image jpg1 = iTextSharp.text.Image.GetInstance(imageURL);
                jpg1.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                jpg1.SpacingBefore = 0f;
                jpg1.SpacingAfter = 0f;
                jpg1.Alignment = Element.ALIGN_CENTER;                
                doc.Add(jpg1);
                // Page Two 
                if (Globals.back != "") 
                {
                    iTextSharp.text.Image jpg2 = iTextSharp.text.Image.GetInstance(folder2 + Globals.back);
                    jpg2.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                    jpg2.SpacingBefore = 0f;
                    jpg2.SpacingAfter = 0f;
                    jpg2.Alignment = Element.ALIGN_CENTER;
                    doc.Add(jpg2);
                }
                // Page Three
                if (Globals.page3 != "")
                {
                    iTextSharp.text.Image jpg3 = iTextSharp.text.Image.GetInstance(folder2 + Globals.page3);
                    jpg3.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                    jpg3.SpacingBefore = 0f;
                    jpg3.SpacingAfter = 0f;
                    jpg3.Alignment = Element.ALIGN_CENTER;
                    doc.Add(jpg3);
                }
                // Page Four
                if (Globals.page4 != "")
                {
                    iTextSharp.text.Image jpg4 = iTextSharp.text.Image.GetInstance(folder2 + Globals.page4);
                    jpg4.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
                    jpg4.SpacingBefore = 0f;
                    jpg4.SpacingAfter = 0f;
                    jpg4.Alignment = Element.ALIGN_CENTER;
                    doc.Add(jpg4);
                }
                doc.Close();

                var attachmentFiles= "";            
                if (File.Exists(docFilePath + filename + ".doc"))
                {
                    attachmentFiles = docFilePath + filename + ".doc;";
                }
                if (File.Exists(pdfFilePath + filename + ".pdf"))
                {
                    attachmentFiles = attachmentFiles + pdfFilePath + filename + ".pdf";
                }

            //Console.WriteLine(attachmentFiles);
            //Console.ReadLine();

                  if (File.Exists(pdfFilePath + filename + ".pdf") && File.Exists(docFilePath + filename + ".doc"))
                  {
                    //Email PDF here                                        
                    using (SqlConnection connection = new SqlConnection("Data Source=SQLCLUSTER\\SQLCLUSTER;Initial Catalog=Stratadial;Integrated Security=SSPI"))
                    {
                        SqlCommand cmd = new SqlCommand("JB_email_update_variable_sender2", connection);
                        cmd.Connection.Open();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@display_name", "Jess Bautista"));
                        cmd.Parameters.Add(new SqlParameter("@email_from", "jesusbautista@medicainc.com"));
                        cmd.ExecuteReader();
                        cmd.Dispose();
                    }

                    using (SqlConnection connection = new SqlConnection("Data Source=SQLCLUSTER\\SQLCLUSTER;Initial Catalog=Stratadial;Integrated Security=SSPI"))
                    {
                        SqlCommand cmd = new SqlCommand("JB_email_generator_generic", connection);
                        cmd.Connection.Open();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@themessage", "<br>Please find attached invite for <b> " + Globals.description + " <b>"));
                        if (Parameters.staticRun)
                        {
                            cmd.Parameters.Add(new SqlParameter("@themessageto", "jesusbautista@medicainc.com"));
                        }
                        else
                        {
                            cmd.Parameters.Add(new SqlParameter("@themessageto", mplanner_email));
                        }
                        cmd.Parameters.Add(new SqlParameter("@thesubject", Globals.description));
                        cmd.Parameters.Add(new SqlParameter("@thefileAttachments", attachmentFiles));
                        cmd.Parameters.Add(new SqlParameter("@theserver", "Nah"));
                        cmd.ExecuteReader();
                        cmd.Dispose();
                    }

                    //Update tracker table to flag it that email was sent
                    SqlConnection sqlConn = new SqlConnection("Data Source=SQLCLUSTER\\SQLCLUSTER;Initial Catalog=Stratadial;Integrated Security=SSPI");
                    sqlConn.Open();
                    SqlCommand cmd1 = new SqlCommand("UPDATE JB_INVITE_TRACKER SET INVITESENT='True' WHERE [NFP Code/Janssen ID]='" + Globals.NFP_Code + "' AND [Apex Job Number]='" + Globals.jobNo + "'", sqlConn);
                    cmd1.ExecuteNonQuery();
                    sqlConn.Close();
                }
                else
                    Console.WriteLine("No file");                
           //catch {  }
        }
    }
    
}
