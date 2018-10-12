using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PDF_Invites
{
    public class test
    {
        public static void MainTest()
        {
            var path = "\\\\fscluster\\webdata\\UPLOADED_FILES\\Medica\\Invitations\\App\\";

            List<Program> prg = null;
            using (StreamReader r = new StreamReader(path + "Data\\program2.json"))
            {
                string json = r.ReadToEnd();
                prg = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Program>>(json);
            }

            foreach (var item in prg)
            {
                //if (item.topic == "Rheum 2016 Simponi Aria Slide Deck")
                //{
                //    foreach (var obj in item.myPages)
                //    {
                //        Console.WriteLine(obj.fname);
                //        //Console.ReadLine();
                //    }
                //}
            }

            //Console.WriteLine("hey");

            var x = "SELECT topic, SHW_ACJANS28_MEETINGPLANNER.email as mplanner_email, t1.[Meeting Coordinator]," +
                         "isnull([apex job number],'Job Number') as [Apex Job Number], " +
                         "isnull([confirmedspeaker], 'NO CONFIRMED SPEAKER') as [Confirmed Speaker]," +
                         " isnull(t2.[credential 1], '') as [Credential 1], " +
                         " isnull(t2.[credential 2], '') as [Credential 2], " +
                         " isnull(t2.[credential 3], '') as [Credential 3], " +
                         " isnull(datename(weekday, convert(datetime, [confirmed event date]))+', '+ " +
                         " datename(month, convert(datetime, [confirmed event date]))+' '+ " +
                         " datename(day, convert(datetime, [confirmed event date]))+', '+ " +
                         " datename(Year, convert(datetime, [confirmed event date]))+' ','No Confirmed Date') as [CED], " +
                         //" isnull([time], 'arrival time')+' '+isnull([Zone], '') as [Arrival Time], " +
                         " isnull([time], 'arrival time') as [Arrival Time], " +
                         " isnull([Presentation time], 'No Presentation Time') as [Presentation Time], " +
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
                         " and t1.[apex job number]=t2.job " +
                         " LEFT OUTER JOIN SHW_ACJANS28_MEETINGPLANNER ON t1.[Meeting Coordinator] = SHW_ACJANS28_MEETINGPLANNER.MeetingPlanner " +
                         " where [apex job number]='' and [nfp code/janssen id]=''";
            Console.WriteLine(x);
            Console.ReadLine();
        }
    }
}
