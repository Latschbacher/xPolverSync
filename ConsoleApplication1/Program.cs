using ConsoleApplication1;
using System;

using System.Windows.Forms;

public class Program
 {

    static void Main(string[] args)
    {

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        if (args.Length == 0)
            Application.Run(new Form1());
        else
        {
            Form1 f = new Form1();
            switch (args[0])
            {
                case "-go": f.btnStart_Click(null, null); break;    //all
                case "-as": f.getAssortments(null, null); break;    //Assortments
                case "-at": f.getAttachments(null, null); break;    //Attachments
                case "-ac": f.accPolter(null, null); break;         //Accept Polter
                case "-r": f.abroPolter(null, null); break;         //Reject(Abbrogate) Polter
                case "-s": f.setPolter(null, null); break;          //SetPolter
                case "-g": f.getPolter(null, null); break;          //GetPolter
                case "-c": f.getContacts(null, null); break;        //GetContacts
            }
        }
    }
}
