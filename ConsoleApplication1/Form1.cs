using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConsoleApplication1.XPPRD;
using ConsoleApplication1.WfpNetService;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Data.SqlClient;
using System.IO;
using System.ServiceModel;
using System.Security;
using System.Runtime.Serialization.Formatters.Binary;

namespace ConsoleApplication1
{
    public partial class Form1 : Form
    {
        Dictionary<string, int> dictorgha = new Dictionary<string, int>(); //"Text" von Polver -> ID von NL
        Dictionary<string, int> dictorghs = new Dictionary<string, int>();
        Dictionary<string, int> dictorggkl = new Dictionary<string, int>();
        Dictionary<int, String> dictnlha = new Dictionary<int, String>(); //sortCode von Polver -> KURZ von NL
        Dictionary<int, String> dictnlhs = new Dictionary<int, String>();
        Dictionary<int, String> dictnlgkl = new Dictionary<int, String>();
        Dictionary<String, String> dictcontacts = new Dictionary<string, string>(); //RecFrom von Polver -> 
        Dictionary<int, String> dictdm = new Dictionary<int, string>(); //Durchmesser von Polter
        Dictionary<int, String> dictlen = new Dictionary<int, string>(); //Länge von Polter

        string text;
        const int SHA256_DIGEST_LENGTH = 32;     

        // Einstellungen setzen
        public Form1()
        {
            InitializeComponent();
            txtb_dictha.Text = Settings.Default.dictionaryha;
            txtb_dicths.Text = Settings.Default.dictionaryhs;
            txtb_dictgkl.Text = Settings.Default.dictionarygkl;
            txtb_service.Text = Settings.Default.webservice;
            txtb_clientendpoint.Text = Settings.Default.client_endpoint;
            txtb_username.Text = Settings.Default.uebergabewebfirma;
            txtb_password.Text = Settings.Default.webfirma;
            txtb_pxinterfversion.Text = Settings.Default.polverinterfaceversion;
            txtb_pxusername.Text = Settings.Default.polverusername;
            txtb_pxpw.Text = Settings.Default.polverpassword;
            txtb_logfile.Text = Settings.Default.logs;
            txtb_contacts.Text = Settings.Default.dictionarycontacts;
            txtb_mailfrom.Text = Settings.Default.mail_from;
            txtb_mailto.Text = Settings.Default.mail_to;
            txtb_smtphost.Text = Settings.Default.smtp_host;
            txtb_smtpport.Text = Settings.Default.smtp_port;
            txtb_smtpuser.Text = Settings.Default.smtp_user;
            txtb_smtppw.Text = Settings.Default.smtp_pw;
            txtb_tmpdocs.Text = Settings.Default.pathtmpdocs;
            txtb_smtppw.PasswordChar = '*';

        }


        public Boolean getPolter_code()
        {
            Boolean asso_ok = false, cont_ok = false, accp_ok = false;

            progressBar1.Value = 0;
            asso_ok = getAssortments_code(); //Wird auf false gesetzt wenn Dictonary erstellen fehlschlägt
            progressBar1.Value = 25;
            cont_ok = getContacts_code(); //Wird auf false gesetzt wenn Dictonary erstellen fehlschlägt
            progressBar1.Value = 50;

            //Wenn Sortimente und Kontakte Okay sind, dann können Polter angenommen werden
            if (asso_ok & cont_ok) 
            {
                accp_ok = accPolter_code();
            }

            if (accp_ok) //Trifft nur zu, wenn Polter & Kontakte erstellt werden konnten
            {
                //NL
                ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
                sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);

                //Variablen
                DataSet dstest = new DataSet();
                string spec, gkl, type, recfromfirma, len, dm, recfromuser;
                String modprivuser = Settings.Default.webfirma;
                StringBuilder sqlstr, sqlc;
                Boolean bimport;

                sqlc = new StringBuilder();


                //Polver Session
                string text = "Start Import PX Polter " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                Session s = new Session();

                XPPRDClient client = new XPPRDClient();
                client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
                CreateSessionRequest csr = new CreateSessionRequest();

                csr.application = 2101;
                csr.clientName = Settings.Default.polverusername;
                csr.interfVersion = Settings.Default.polverinterfaceversion;
                CreateSessionResponse csrp = client.CreateSession(csr);
                s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
                s.sessionID = csrp.sessionID;
                text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Polver Polter Request
                GetXPViewRequest req = new GetXPViewRequest(s, 1); //viewType = 1 -> Eigene Polter //viewType = 2 -> Angebotene Polter, diese werden im vorhinein akzeptiert
                GetXPViewResponse pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                text = "Import Polter in NL " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);


                sqlstr = new StringBuilder("select * From BENUTZERWEBFIRMA where ID_WEBFIRMA = (select id from WEBFIRMA where firma = '" + Settings.Default.uebergabewebfirma + "')");
                dstest = sc.SelectSQL(sqlstr.ToString(), "sens");


                foreach (PolterResp p in pr.polters)
                {

                    if (progressBar1.Value < 75)
                    {
                        progressBar1.Value = progressBar1.Value + 1;
                    }
                    try
                    {
                        //Wenn Fremdschlüssel aus NL gefüllt ist
                        if (p.foreignRef == "")
                        {
                            //Polterfilter
                            if (txtb_EinzelnerPolter.Text != "")
                            {
                                if (p.key.polterID.ToString() == txtb_EinzelnerPolter.Text)
                                {
                                    bimport = false;
                                    //TryGetValue setzt Ausgabewert auf null falls kein Wert gefunden wird
                                    // Im Vorhinein auf "" setzen, falls TryGetValue kein Ergebnis findet
                                    //spec = "";
                                    //type = "";
                                    //gkl = "";
                                    //
                                    dictnlha.TryGetValue(p.sortCode, out spec);
                                    dictnlhs.TryGetValue(p.sortCode, out type);
                                    dictnlgkl.TryGetValue(p.sortCode, out gkl);
                                    dictdm.TryGetValue(p.sortCode, out dm);
                                    dictlen.TryGetValue(p.sortCode, out len);

                                    //Koordinaten
                                    float xkoord, ykoord;
                                    float.TryParse(p.position.posN, out xkoord);
                                    float.TryParse(p.position.posE, out ykoord);
                                    xkoord = xkoord / 1000000;
                                    ykoord = ykoord / 1000000;

                                    //Recfrom standardmäßig von Benutzer Unknown aus der Webfirma ZHH wenn nicht gefüllt
                                    if (p.recFrom != null)
                                    {
                                        recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                        recfromfirma = "";
                                        dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                        if (recfromfirma != null)
                                        {
                                            bimport = true; //Falls Firma angegeben wurde, aber nicht mittels Dictionary gefunden werden kann, wird er beim Import übersprungen
                                        }
                                    }
                                    else
                                    {
                                        recfromuser = "Unknown";
                                        recfromfirma = "Unknown";
                                        bimport = true;
                                    }
                                    if (bimport == true)
                                    {
                                        if (recfromfirma == Settings.Default.uebergabewebfirma | recfromfirma == "Unknown")
                                        {

                                            //sqlstr = new StringBuilder("select * From BENUTZERWEBFIRMA where ID_WEBFIRMA = (select id from WEBFIRMA where firma = '" + Settings.Default.uebergabewebfirma + "') AND BENUTZERNAME = '" + recfromuser + "'");

                                            //dstest = sc.SelectSQL(sqlstr.ToString(), "sens");

                                            DataRow[] foundContact = dstest.Tables[0].Select("BENUTZERNAME = '" + recfromuser + "'");

                                            //Polter kommt von xPolver-ZHH
                                            //Abprüfen ob Benutzername schon existiert
                                            if (foundContact.Length == 0)
                                            {
                                                try
                                                {
                                                    //EMAIL
                                                    MailMessage mail = new MailMessage(Settings.Default.mail_from, Settings.Default.mail_to);
                                                    SmtpClient smtpclient = new SmtpClient();
                                                    //System.Net.NetworkCredential NC = new System.Net.NetworkCredential("qms@winforstpro.com", "netlog123");
                                                    System.Net.NetworkCredential NC = new System.Net.NetworkCredential(Settings.Default.smtp_user, Settings.Default.smtp_pw);

                                                    smtpclient.UseDefaultCredentials = false;
                                                    smtpclient.Port = int.Parse(Settings.Default.smtp_port);
                                                    smtpclient.Host = Settings.Default.smtp_host;
                                                    //smtpclient.Host = "smtp.1und1.de";
                                                    smtpclient.Credentials = NC;
                                                    mail.Subject = "xPolver Schnittstelle Adresszuordnung"; //
                                                    mail.Body = "Adressen Dictionary prüfen!";
                                                    smtpclient.Send(mail);
                                                }
                                                catch
                                                {
                                                    //MessageBox.Show("Es ist ein Problem beim Emailversand aufgetreten!");
                                                }
                                            }


                                            //Benutzername in xPolver-ZHH existiert bereits
                                            else
                                            {

                                                //Import Polter und Übergabe
                                                sqlstr = new StringBuilder("if not exists (select Polter.ID from polter left join WEBFIRMA on WEBFIRMA.ID = polter.ID_WEBFIRMA where polter.Nummer = 'PO" + p.key.polterID + "' AND WEBFIRMA.FIRMA = '" + Settings.Default.uebergabewebfirma + "')" +
                                                                                          " begin" +
                                                                                          " insert into polter (NUMMER, KOORDINATE_X, KOORDINATE_Y, WGS84_X, WGS84_Y, ID_WEBFIRMA, BEMERKUNG, los)" +
                                                                                          " VALUES ('PO" + p.key.polterID + "', '" + ykoord + "', '" + xkoord + "', replace('" + ykoord + "',',','.'), replace('" + xkoord + "',',','.'), (select ID from WEBFIRMA where FIRMA = ('" + Settings.Default.uebergabewebfirma + "')),left('" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "', 147) + '...', '" + p.ticketID.ToString() + "')" +
                                                                                          " insert into POLTERPOOL (ID_POLTER, ID_WEBFIRMA_POOL, ID_BENUTZERWEBFIRMA, DATUM, STUCK, KUBATUR, ID_T_HOLZART, ID_T_HOLZSORTE, ID_T_GUETEKLASSE, ID_ADR_WALD, MODIFY_DATE, id_t_einheit, cosemat)" +
                                                                                          " VALUES ((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "')), (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), (select ID from BENUTZERWEBFIRMA where BENUTZERNAME = '" + recfromuser + "' and id_webfirma = (select id from webfirma where firma = '" + Settings.Default.uebergabewebfirma + "')), '" + p.initPubDate + "', " + p.plantQuantity + ", " + p.size + ", " +
                                                                                          " (select id from t_holzart where kurz = '" + spec + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_holzsorte where kurz = '" + type + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_gueteklasse where kurz = '" + gkl + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), Nullif('" + p.ownerRef + "',''), getdate(), 1, 1) " +
                                                                                          " insert into POLTERUEBERGABE (ID_POLTER, ID_WEBFIRMA, DATUM, ID_WEBFIRMA_ERRSTELLER)" +
                                                                                          " VALUES((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))," +
                                                                                          " (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), GETDATE(),(select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))" +
                                                                                          " end else begin" +

                                                                                          " update polter set nummer = 'PO" + p.key.polterID + "'" +
                                                                                          ", KOORDINATE_X = '" + ykoord + "'" +
                                                                                          ", KOORDINATE_Y = '" + xkoord + "'" +
                                                                                          ", WGS84_X = replace('" + ykoord + "',',','.')" +
                                                                                          ", WGS84_Y = replace('" + xkoord + "',',','.')" +
                                                                                          //", Bemerkung = '" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "' " +
                                                                                          " where polter.id = (select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))" +
                                                                                          " end ");

                                                sqlc.Append(sqlstr.ToString());
                                                //sc.SelectSQL(sqlstr.ToString(), "sens");
                                            }


                                        }
                                        //Polter kommt von Firma die es in NL gibt und im Dictonary gefunden wurde
                                        else
                                        {
                                            sqlstr = new StringBuilder("if not exists (select Polter.ID from polter left join WEBFIRMA on WEBFIRMA.ID = polter.ID_WEBFIRMA where polter.Nummer = 'PO" + p.key.polterID + "' AND WEBFIRMA.FIRMA = '" + recfromfirma + "')" +
                                                                  " begin" +
                                                                  " insert into polter (NUMMER, KOORDINATE_X, KOORDINATE_Y, WGS84_X, WGS84_Y, ID_WEBFIRMA, BEMERKUNG, los)" +
                                                                  " VALUES ('PO" + p.key.polterID + "', '" + ykoord + "', '" + xkoord + "', replace('" + ykoord + "',',','.'), replace('" + xkoord + "',',','.'), (select ID from WEBFIRMA where FIRMA = ('" + recfromfirma + "')),left('" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "', 147) + '...', '" + p.ticketID.ToString() + "')" +
                                                                  " insert into POLTERPOOL (ID_POLTER, ID_WEBFIRMA_POOL, ID_BENUTZERWEBFIRMA, DATUM, STUCK, KUBATUR, ID_T_HOLZART, ID_T_HOLZSORTE, ID_T_GUETEKLASSE, ID_ADR_WALD, MODIFY_DATE, id_t_einheit, cosemat)" +
                                                                  " VALUES ((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')), (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), (select top 1 ID from BENUTZERWEBFIRMA where id_webfirma = (select id from webfirma where firma = '" + recfromfirma + "')), '" + p.initPubDate + "', " + p.plantQuantity + ", " + p.size + ", " +
                                                                  " (select id from t_holzart where kurz = '" + spec + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_holzsorte where kurz = '" + type + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_gueteklasse where kurz = '" + gkl + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), Nullif('" + p.ownerRef + "',''), getdate(), 1, 1)" +
                                                                  " insert into POLTERUEBERGABE (ID_POLTER, ID_WEBFIRMA, DATUM, ID_WEBFIRMA_ERRSTELLER)" +
                                                                  " VALUES((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))," +
                                                                  " (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), GETDATE(),(select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))" +

                                                                  " end else begin" +

                                                                  " update polter set nummer = 'PO" + p.key.polterID + "'" +
                                                                  ", KOORDINATE_X = '" + ykoord + "'" +
                                                                  ", KOORDINATE_Y = '" + xkoord + "'" +
                                                                  ", WGS84_X = replace('" + ykoord + "',',','.')" +
                                                                  ", WGS84_Y = replace('" + xkoord + "',',','.')" +
                                                                  //", Bemerkung = '" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "' " +
                                                                  " where polter.id = (select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))" +
                                                                  " end ");

                                            sqlc.Append(sqlstr.ToString());
                                            //sc.SelectSQL(sqlstr.ToString(), "sens");
                                        }
                                    }
                                }

                            }

                            //Polterfilter ENDE
                            else
                            {
                                bimport = false;
                                //TryGetValue setzt Ausgabewert auf null falls kein Wert gefunden wird
                                // Im Vorhinein auf "" setzen, falls TryGetValue kein Ergebnis findet
                                //spec = "";
                                //type = "";
                                //gkl = "";
                                //
                                dictnlha.TryGetValue(p.sortCode, out spec);
                                dictnlhs.TryGetValue(p.sortCode, out type);
                                dictnlgkl.TryGetValue(p.sortCode, out gkl);
                                dictdm.TryGetValue(p.sortCode, out dm);
                                dictlen.TryGetValue(p.sortCode, out len);

                                //Koordinaten
                                float xkoord, ykoord;
                                float.TryParse(p.position.posN, out xkoord);
                                float.TryParse(p.position.posE, out ykoord);
                                xkoord = xkoord / 1000000;
                                ykoord = ykoord / 1000000;

                                //Recfrom standardmäßig von Benutzer Unknown aus der Webfirma ZHH wenn nicht gefüllt
                                if (p.recFrom != null)
                                {
                                    recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                    recfromfirma = "";
                                    dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                    if (recfromfirma != null)
                                    {
                                        bimport = true; //Falls Firma angegeben wurde, aber nicht mittels Dictionary gefunden werden kann, wird er beim Import übersprungen
                                    }
                                }
                                else
                                {
                                    recfromuser = "Unknown";
                                    recfromfirma = "Unknown";
                                    bimport = true;
                                }
                                if (bimport == true)
                                {
                                    if (recfromfirma == Settings.Default.uebergabewebfirma | recfromfirma == "Unknown")
                                    {

                                        //sqlstr = new StringBuilder("select * From BENUTZERWEBFIRMA where ID_WEBFIRMA = (select id from WEBFIRMA where firma = '" + Settings.Default.uebergabewebfirma + "') AND BENUTZERNAME = '" + recfromuser + "'");

                                        //dstest = sc.SelectSQL(sqlstr.ToString(), "sens");

                                        DataRow[] foundContact = dstest.Tables[0].Select("BENUTZERNAME = '" + recfromuser + "'");

                                        //Polter kommt von xPolver-ZHH
                                        //Abprüfen ob Benutzername schon existiert
                                        if (foundContact.Length == 0)
                                        {
                                            try
                                            {
                                                //EMAIL
                                                MailMessage mail = new MailMessage(Settings.Default.mail_from, Settings.Default.mail_to);
                                                SmtpClient smtpclient = new SmtpClient();
                                                //System.Net.NetworkCredential NC = new System.Net.NetworkCredential("qms@winforstpro.com", "netlog123");
                                                System.Net.NetworkCredential NC = new System.Net.NetworkCredential(Settings.Default.smtp_user, Settings.Default.smtp_pw);

                                                smtpclient.UseDefaultCredentials = false;
                                                smtpclient.Port = int.Parse(Settings.Default.smtp_port);
                                                smtpclient.Host = Settings.Default.smtp_host;
                                                //smtpclient.Host = "smtp.1und1.de";
                                                smtpclient.Credentials = NC;
                                                mail.Subject = "xPolver Schnittstelle Adresszuordnung"; //
                                                mail.Body = "Adressen Dictionary prüfen!";
                                                smtpclient.Send(mail);
                                            }
                                            catch
                                            {
                                                //MessageBox.Show("Es ist ein Problem beim Emailversand aufgetreten!");
                                            }
                                        }
                                        //Benutzername in xPolver-ZHH existiert bereits
                                        else
                                        {

                                            //Import Polter und Übergabe
                                            sqlstr = new StringBuilder("if not exists (select Polter.ID from polter left join WEBFIRMA on WEBFIRMA.ID = polter.ID_WEBFIRMA where polter.Nummer = 'PO" + p.key.polterID + "' AND WEBFIRMA.FIRMA = '" + Settings.Default.uebergabewebfirma + "')" +
                                                                                      " begin" +
                                                                                      " insert into polter (NUMMER, KOORDINATE_X, KOORDINATE_Y, WGS84_X, WGS84_Y, ID_WEBFIRMA, BEMERKUNG, los)" +
                                                                                      " VALUES ('PO" + p.key.polterID + "', '" + ykoord + "', '" + xkoord + "', replace('" + ykoord + "',',','.'), replace('" + xkoord + "',',','.'), (select ID from WEBFIRMA where FIRMA = ('" + Settings.Default.uebergabewebfirma + "')),left('" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "', 147) + '...', '" + p.ticketID.ToString() + "')" +
                                                                                      " insert into POLTERPOOL (ID_POLTER, ID_WEBFIRMA_POOL, ID_BENUTZERWEBFIRMA, DATUM, STUCK, KUBATUR, ID_T_HOLZART, ID_T_HOLZSORTE, ID_T_GUETEKLASSE, ID_ADR_WALD, MODIFY_DATE, id_t_einheit,cosemat)" +
                                                                                      " VALUES ((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "')), (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), (select ID from BENUTZERWEBFIRMA where BENUTZERNAME = '" + recfromuser + "' and id_webfirma = (select id from webfirma where firma = '" + Settings.Default.uebergabewebfirma + "')), '" + p.initPubDate + "', " + p.plantQuantity + ", " + p.size + ", " +
                                                                                      " (select id from t_holzart where kurz = '" + spec + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_holzsorte where kurz = '" + type + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_gueteklasse where kurz = '" + gkl + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), Nullif('" + p.ownerRef + "',''), getdate(), 1, 1)" +
                                                                                      " insert into POLTERUEBERGABE (ID_POLTER, ID_WEBFIRMA, DATUM, ID_WEBFIRMA_ERRSTELLER)" +
                                                                                      " VALUES((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))," +
                                                                                      " (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), GETDATE(),(select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))" +
                                                                                      " end else begin" +

                                                                                      " update polter set nummer = 'PO" + p.key.polterID + "'" +
                                                                                      ", KOORDINATE_X = '" + ykoord + "'" +
                                                                                      ", KOORDINATE_Y = '" + xkoord + "'" +
                                                                                      ", WGS84_X = replace('" + ykoord + "',',','.')" +
                                                                                      ", WGS84_Y = replace('" + xkoord + "',',','.')" +
                                                                                      //", Bemerkung = '" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "' " +
                                                                                      " where polter.id = (select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "'))" +
                                                                                      " end ");

                                            sqlc.Append(sqlstr.ToString());
                                            //sc.SelectSQL(sqlstr.ToString(), "sens");
                                        }


                                    }
                                    //Polter kommt von Firma die es in NL gibt und im Dictonary gefunden wurde
                                    else
                                    {
                                        sqlstr = new StringBuilder("if not exists (select Polter.ID from polter left join WEBFIRMA on WEBFIRMA.ID = polter.ID_WEBFIRMA where polter.Nummer = 'PO" + p.key.polterID + "' AND WEBFIRMA.FIRMA = '" + recfromfirma + "')" +
                                                              " begin" +
                                                              " insert into polter (NUMMER, KOORDINATE_X, KOORDINATE_Y, WGS84_X, WGS84_Y, ID_WEBFIRMA, BEMERKUNG, los)" +
                                                              " VALUES ('PO" + p.key.polterID + "', '" + ykoord + "', '" + xkoord + "', replace('" + ykoord + "',',','.'), replace('" + xkoord + "',',','.'), (select ID from WEBFIRMA where FIRMA = ('" + recfromfirma + "')),left('" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "', 147) + '...', '" + p.ticketID.ToString() + "')" +
                                                              " insert into POLTERPOOL (ID_POLTER, ID_WEBFIRMA_POOL, ID_BENUTZERWEBFIRMA, DATUM, STUCK, KUBATUR, ID_T_HOLZART, ID_T_HOLZSORTE, ID_T_GUETEKLASSE, ID_ADR_WALD, MODIFY_DATE, id_t_einheit, cosemat)" +
                                                              " VALUES ((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')), (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), (select top 1 ID from BENUTZERWEBFIRMA where id_webfirma = (select id from webfirma where firma = '" + recfromfirma + "')), '" + p.initPubDate + "', " + p.plantQuantity + ", " + p.size + ", " +
                                                              " (select id from t_holzart where kurz = '" + spec + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_holzsorte where kurz = '" + type + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), (select id from t_gueteklasse where kurz = '" + gkl + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')), Nullif('" + p.ownerRef + "',''), getdate(), 1, 1)" +
                                                              " insert into POLTERUEBERGABE (ID_POLTER, ID_WEBFIRMA, DATUM, ID_WEBFIRMA_ERRSTELLER)" +
                                                              " VALUES((select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))," +
                                                              " (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "'), GETDATE(),(select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))" +

                                                              " end else begin" +

                                                              " update polter set nummer = 'PO" + p.key.polterID + "'" +
                                                              ", KOORDINATE_X = '" + ykoord + "'" +
                                                              ", KOORDINATE_Y = '" + xkoord + "'" +
                                                              ", WGS84_X = replace('" + ykoord + "',',','.')" +
                                                              ", WGS84_Y = replace('" + xkoord + "',',','.')" +
                                                              //", Bemerkung = '" + p.polterInfo + "; " + p.modPrivUser + "; " + p.ownerInfo + "; SortCode: " + p.sortCode + ", " + spec + ", " + type + ", " + gkl + ", " + len + ", " + dm + "' " +
                                                              " where polter.id = (select ID from POLTER where NUMMER = 'PO" + p.key.polterID + "' AND ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "'))" +
                                                              " end ");

                                        sqlc.Append(sqlstr.ToString());
                                        //sc.SelectSQL(sqlstr.ToString(), "sens");
                                    }
                                }
                            }

                        }
                        //client.RenewSession(new RenewSessionRequest(s));

                        //Hier Fremdschlüssel bei Polter setzen


                    }
                    catch
                    {
                        this.text = "Fehler Polter Import Polternr.:PO" + p.key.polterID + " " + DateTime.Now.TimeOfDay + "\n";
                        System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                    }

                }

                if (sqlc.ToString().Length > 0)
                {
                    sc.SelectSQL(sqlc.ToString(), "sens");
                }

                sqlstr = new StringBuilder("Select ID as PID, NUMMER as POLTERNR from POLTER where POLTER.NUMMER in (  ");

                foreach (PolterResp p in pr.polters)
                {
                    if (p.foreignRef == "")
                    {
                        sqlstr.Append("'PO" + p.key.polterID + "',");
                    }
                }
                sqlstr.Length--;
                sqlstr.Append(")");

                DataSet dsPIDPolver = new DataSet();
                dsPIDPolver = sc.SelectSQL(sqlstr.ToString(), "sens");
                DataTable dtPIDPolver = new DataTable();
                dtPIDPolver.Columns.Add("PID");
                dtPIDPolver.Columns.Add("POLTERNR");
                //ToDo bei keinen Poltern mit leeren Fremdschlüsseln
                if (dsPIDPolver != null)
                {
                    dtPIDPolver = dsPIDPolver.Tables["OUT"];
                    dtPIDPolver.PrimaryKey = new DataColumn[] { dtPIDPolver.Columns["POLTERNR"] };


                    DataRowCollection drc = dtPIDPolver.Rows;
                    XPPRD.PolterReq updatepolterrequest = new XPPRD.PolterReq();

                    foreach (PolterResp p in pr.polters)
                    {
                        if (p.foreignRef.ToString() == "")
                        {
                            if (drc.Contains("PO" + p.key.polterID.ToString()))
                            {
                                updatepolterrequest.addInfo1 = p.addInfo1;
                                updatepolterrequest.addInfo2 = p.addInfo2;
                                updatepolterrequest.addInfo3 = p.addInfo3;
                                updatepolterrequest.addInfo4 = p.addInfo4;
                                updatepolterrequest.colorKey = p.colorKey;
                                updatepolterrequest.colorKeySpecified = true;
                                updatepolterrequest.orgInfo1 = p.orgInfo1;
                                updatepolterrequest.orgInfo2 = p.orgInfo2;
                                updatepolterrequest.ownerCode = p.ownerCode;
                                updatepolterrequest.ownerCodeSpecified = true;
                                updatepolterrequest.ownerInfo = p.ownerInfo;
                                updatepolterrequest.ownerRef = p.ownerRef;
                                updatepolterrequest.size = Int32.Parse(p.size.ToString());
                                updatepolterrequest.sizeSpecified = true;
                                updatepolterrequest.sortCode = p.sortCode;
                                updatepolterrequest.sortCodeSpecified = true;
                                updatepolterrequest.ticketID = p.ticketID;

                                updatepolterrequest.polterInfo = p.polterInfo;
                                updatepolterrequest.foreignRef = drc.Find("PO" + p.key.polterID.ToString()).ItemArray[0].ToString();

                                UpdateXPolterRequest upr = new UpdateXPolterRequest(s, p.key, p.position, updatepolterrequest);
                                UpdateXPolterResponse uprr = client.UpdateXPolter(upr);
                            }
                        }
                    }
                }

                sc.Close();
                client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                progressBar1.Value = 100;
                return true;
            }
            else
            {
                progressBar1.Value = 100;
                return false;
            }
        }

        // Restmenge von Poltern zurückschreiben und Polter als abgefahren melden wenn Restmenge = 0
        public Boolean setPolter_code()
        {
            Boolean cont_ok = getContacts_code(); //Kontakte-Dictionary neu erstellen

            if (cont_ok) //Falls Kontakte OK, dann Rückschreiben von Mengen möglich
            {

                //NL
                ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
                sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);

                //Variablen
                DataSet ds1 = new DataSet();
                DataSet ds2 = new DataSet();
                String recfromuser, recfromfirma;
                float number;
                StringBuilder sqlstr;
                Int16 af = 1;

                //Polver Session
                this.text = "Start SetPolter " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);


                text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                Session s = new Session();
                XPPRDClient client = new XPPRDClient();
                client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
                CreateSessionRequest csr = new CreateSessionRequest();


                csr.application = 2101;
                csr.clientName = Settings.Default.polverusername;
                csr.interfVersion = Settings.Default.polverinterfaceversion;
                CreateSessionResponse csrp = client.CreateSession(csr);
                s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
                s.sessionID = csrp.sessionID;
                text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Polver Polter Request
                GetXPViewRequest req = new GetXPViewRequest(s, 1);
                GetXPViewResponse pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Prüfen ob Authentifizierung funktioniert hat
                if (pr.fb.code == 0)
                {
                    af = 0;
                }

                //Restmenge des Polters zurückschreiben
                //Polter als abgefahren markieren wenn Rest_Kubatur = 0

                foreach (PolterResp p in pr.polters)
                {

                    try
                    {
                        //Polterfilter
                        if (txtb_EinzelnerPolter.Text != "")
                        {
                            if (p.key.polterID.ToString() == txtb_EinzelnerPolter.Text)
                            {
                                recfromfirma = "";

                                if (p.recFrom != null)
                                {
                                    recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                    //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                    dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                }
                                else
                                {
                                    recfromuser = "Unknown";
                                    recfromfirma = "Unknown";
                                }

                                //Firma ist bekannt
                                if (recfromfirma != "Unknown" & recfromfirma != null)
                                {
                                    sqlstr = new StringBuilder("select top 1 rest_kubatur from AUFTRAG_TRANSPORT where ID_POLTER = (select id from polter where nummer = 'PO" + p.key.polterID + "' " +
                                        "  and ID_WEBFIRMA = (select id from webfirma where firma = '" + recfromfirma + "')) order by ABFUHRDATUM DESC, ID DESC");
                                    ds2 = sc.SelectSQL(sqlstr.ToString(), "sens");
                                    if (ds2 != null)
                                    {

                                        if (ds2.Tables[0].Rows.Count != 0)
                                        {
                                            SetSizeXPolterRequest ssxpr = new SetSizeXPolterRequest(s, p.key, 2001, ds2.Tables[0].Rows[0][0].ToString(), "Ist kleiner");
                                            SetSizeXPolterResponse ssxprr = client.SetSizeXPolter(ssxpr);
                                            text = "Set Size Polter Feedback: " + ssxprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                            System.IO.File.AppendAllText(Settings.Default.logs, text);

                                            float.TryParse((ds2.Tables[0].Rows[0][0].ToString()), out number);

                                            if (number == 0)
                                            {
                                                SetXPProcessedRequest sxppr = new SetXPProcessedRequest(s, p.key, 1);
                                                SetXPProcessedResponse sxpprr = client.SetXPProcessed(sxppr);
                                                text = "Set Processed Polter Feedback: " + sxpprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                                System.IO.File.AppendAllText(Settings.Default.logs, text);
                                            }
                                        }
                                    }
                                }

                                //p.recFrom = null -> Zürichholz hat Polter erstellt
                                if (recfromfirma == "Unknown")
                                {
                                    sqlstr = new StringBuilder("select top 1 rest_kubatur from AUFTRAG_TRANSPORT where ID_POLTER = (select id from polter where nummer = 'PO" + p.key.polterID + "' " +
                                        "  and ID_WEBFIRMA = (select id from webfirma where firma = '" + recfromfirma + "')) order by ABFUHRDATUM DESC, ID DESC");
                                    ds2 = sc.SelectSQL(sqlstr.ToString(), "sens");
                                    if (ds2 != null)
                                    {

                                        if (ds2.Tables[0].Rows.Count != 0)
                                        {
                                            SetSizeXPolterRequest ssxpr = new SetSizeXPolterRequest(s, p.key, 2001, ds2.Tables[0].Rows[0][0].ToString(), "Ist kleiner");
                                            SetSizeXPolterResponse ssxprr = client.SetSizeXPolter(ssxpr);
                                            text = "Set Size Polter Feedback: " + ssxprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                            System.IO.File.AppendAllText(Settings.Default.logs, text);

                                            float.TryParse((ds2.Tables[0].Rows[0][0].ToString()), out number);

                                            if (number == 0)
                                            {
                                                SetXPProcessedRequest sxppr = new SetXPProcessedRequest(s, p.key, 1);
                                                SetXPProcessedResponse sxpprr = client.SetXPProcessed(sxppr);
                                                text = "Set Processed Polter Feedback: " + sxpprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                                System.IO.File.AppendAllText(Settings.Default.logs, text);
                                            }
                                        }
                                    }
                                }
                            }
                        
                        }
                        //Polterfilter ENDE
                        else
                        {
                            recfromfirma = "";

                            if (p.recFrom != null)
                            {
                                recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                            }
                            else
                            {
                                recfromuser = "Unknown";
                                recfromfirma = "Unknown";
                            }

                            //Firma ist bekannt
                            if (recfromfirma != "Unknown" & recfromfirma != null)
                            {
                                sqlstr = new StringBuilder("select top 1 rest_kubatur from AUFTRAG_TRANSPORT where ID_POLTER = (select id from polter where nummer = 'PO" + p.key.polterID + "' " +
                                    "  and ID_WEBFIRMA = (select id from webfirma where firma = '" + recfromfirma + "')) order by ABFUHRDATUM desc, id desc");
                                ds2 = sc.SelectSQL(sqlstr.ToString(), "sens");
                                if (ds2 != null)
                                {

                                    if (ds2.Tables[0].Rows.Count != 0)
                                    {
                                        SetSizeXPolterRequest ssxpr = new SetSizeXPolterRequest(s, p.key, 2001, ds2.Tables[0].Rows[0][0].ToString(), "Ist kleiner");
                                        SetSizeXPolterResponse ssxprr = client.SetSizeXPolter(ssxpr);
                                        text = "Set Size Polter Feedback: " + ssxprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                        System.IO.File.AppendAllText(Settings.Default.logs, text);

                                        float.TryParse((ds2.Tables[0].Rows[0][0].ToString()), out number);

                                        if (number == 0)
                                        {
                                            SetXPProcessedRequest sxppr = new SetXPProcessedRequest(s, p.key, 1);
                                            SetXPProcessedResponse sxpprr = client.SetXPProcessed(sxppr);
                                            text = "Set Processed Polter Feedback: " + sxpprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                            System.IO.File.AppendAllText(Settings.Default.logs, text);
                                        }
                                    }
                                }
                            }

                            //p.recFrom = null -> Zürichholz hat Polter erstellt
                            if (recfromfirma == "Unknown")
                            {
                                sqlstr = new StringBuilder("select top 1 rest_kubatur from AUFTRAG_TRANSPORT where ID_POLTER = (select id from polter where nummer = 'PO" + p.key.polterID + "' " +
                                    "  and ID_WEBFIRMA = (select id from webfirma where firma = '" + Settings.Default.uebergabewebfirma + "')) order by ABFUHRDATUM desc, id desc"); //original settings.default.webfirma
                                ds2 = sc.SelectSQL(sqlstr.ToString(), "sens");
                                if (ds2 != null)
                                {

                                    if (ds2.Tables[0].Rows.Count != 0)
                                    {
                                        SetSizeXPolterRequest ssxpr = new SetSizeXPolterRequest(s, p.key, 2001, ds2.Tables[0].Rows[0][0].ToString(), "Ist kleiner");
                                        SetSizeXPolterResponse ssxprr = client.SetSizeXPolter(ssxpr);
                                        text = "Set Size Polter Feedback: " + ssxprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                        System.IO.File.AppendAllText(Settings.Default.logs, text);

                                        float.TryParse((ds2.Tables[0].Rows[0][0].ToString()), out number);

                                        if (number == 0)
                                        {
                                            SetXPProcessedRequest sxppr = new SetXPProcessedRequest(s, p.key, 1);
                                            SetXPProcessedResponse sxpprr = client.SetXPProcessed(sxppr);
                                            text = "Set Processed Polter Feedback: " + sxpprr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                            System.IO.File.AppendAllText(Settings.Default.logs, text);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        this.text = "Fehler Polter setsize/setprocessed " + DateTime.Now.TimeOfDay + "\n";
                        System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                    }
                    //client.RenewSession(new RenewSessionRequest(s));
                }



                    
                //Rückschreiben der Menge in NL
                sqlstr = new StringBuilder("select p.ID, pp.kubatur, isnull(at.Rest_Kubatur,99999) as Rest_Kubatur, 0 as FIP from POLTER p left join POLTERPOOL pp on p.ID = pp.ID_POLTER left join POLTERUEBERGABE pu on p.ID = pu.ID_POLTER " +
                    " outer apply (Select top 1 at.Rest_Kubatur from AUFTRAG_TRANSPORT at where at.ID_POLTER = p.ID order by at.ABFUHRDATUM desc, at.id desc) at" +
                    " where (p.ID_WEBFIRMA = (select id from WEBFIRMA where firma = '" + Settings.Default.webfirma + "') or pu.ID_WEBFIRMA = (select id from WEBFIRMA where firma = '" + Settings.Default.webfirma + "')) and pp.KUBATUR <> 0 and cosemat = 1");
                DataSet nlpolter = new DataSet();
                nlpolter = sc.SelectSQL(sqlstr.ToString(), "sens");

                sqlstr = new StringBuilder();

                //Alle eigenen Polter durchgehen ob ein Transportauftrag in NL gemacht werden muss
                foreach (DataRow dr in nlpolter.Tables[0].Rows)
                {

                    if (float.Parse(dr["Rest_Kubatur"].ToString()) != 0)
                    {
                        foreach (PolterResp p in pr.polters)
                        {
                            if (p.foreignRef == dr["ID"].ToString())
                            {
                                dr["FIP"] = 1;
                                if (float.Parse(dr["Rest_Kubatur"].ToString()) != float.Parse(p.size) && float.Parse(dr["Kubatur"].ToString()) != float.Parse(p.size))
                                {
                                    recfromfirma = "";

                                    if (p.recFrom != null)
                                    {
                                        recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                        //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                        dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                    }
                                    else
                                    {
                                        recfromuser = "Unknown";
                                        recfromfirma = "Unknown";
                                    }

                                    //Firma ist bekannt
                                    if (recfromfirma != "Unknown" & recfromfirma != null)
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + " , 0, '" + p.size + "', getdate(), 723, newid() ");
                                    }

                                    //p.recFrom = null -> Zürichholz hat Polter erstellt
                                    if (recfromfirma == "Unknown")
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + ", 0, '" + p.size + "', getdate(), 280, newid() ");
                                    }

                                }
                            }
                        }
                    }
                }

                req = new GetXPViewRequest(s, 3);
                pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Alle eigenen Polter durchgehen ob ein Transportauftrag in NL gemacht werden muss
                foreach (DataRow dr in nlpolter.Tables[0].Rows)
                {

                    if (float.Parse(dr["Rest_Kubatur"].ToString()) != 0)
                    {
                        foreach (PolterResp p in pr.polters)
                        {
                            if (p.foreignRef == dr["ID"].ToString())
                            {
                                dr["FIP"] = 1;
                                if (Math.Floor(float.Parse(dr["Rest_Kubatur"].ToString())) != float.Parse(p.size) && Math.Floor(float.Parse(dr["Kubatur"].ToString())) != float.Parse(p.size))
                                {
                                    recfromfirma = "";

                                    if (p.recFrom != null)
                                    {
                                        recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                        //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                        dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                    }
                                    else
                                    {
                                        recfromuser = "Unknown";
                                        recfromfirma = "Unknown";
                                    }

                                    //Firma ist bekannt
                                    if (recfromfirma != "Unknown" & recfromfirma != null)
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + " , 0, '" + p.size + "', getdate(), 723, newid() ");
                                    }

                                    //p.recFrom = null -> Zürichholz hat Polter erstellt
                                    if (recfromfirma == "Unknown")
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + ", 0, '" + p.size + "', getdate(), 280, newid() ");
                                    }
                                }
                            }
                        }
                    }
                }


                req = new GetXPViewRequest(s, 7);
                pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Alle eigenen Polter durchgehen ob ein Transportauftrag in NL gemacht werden muss
                foreach (DataRow dr in nlpolter.Tables[0].Rows)
                {

                    if (float.Parse(dr["Rest_Kubatur"].ToString()) != 0)
                    {
                        foreach (PolterResp p in pr.polters)
                        {
                            if (p.foreignRef == dr["ID"].ToString())
                            {
                                dr["FIP"] = 1;
                                if ((Math.Floor(float.Parse(dr["Rest_Kubatur"].ToString())) != float.Parse(p.size)) && (Math.Floor(float.Parse(dr["Kubatur"].ToString())) != float.Parse(p.size) || Math.Floor(float.Parse(dr["Kubatur"].ToString())) == 0) )
                                {
                                    recfromfirma = "";

                                    if (p.recFrom != null)
                                    {
                                        recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                        //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                        dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                    }
                                    else
                                    {
                                        recfromuser = "Unknown";
                                        recfromfirma = "Unknown";
                                    }

                                    //Firma ist bekannt
                                    if (recfromfirma != "Unknown" & recfromfirma != null)
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + " , 0, '" + p.size + "', getdate(), 723, newid() ");
                                    }

                                    //p.recFrom = null -> Zürichholz hat Polter erstellt
                                    if (recfromfirma == "Unknown")
                                    {
                                        sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                            " select " + p.foreignRef + ", 0, '" + p.size + "', getdate(), 280, newid() ");
                                    }
                                }
                            }
                        }
                    }
                }


                //Polver Polter Request
                req = new GetXPViewRequest(s, 4);
                pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Prüfen ob Authentifizierung funktioniert hat
                if (pr.fb.code == 0 && af == 0)
                {
                    af = 0;
                }
                else
                {
                    af = 1;
                }

                //Nachschauen in Weitergegebenen Poltern
                foreach (DataRow dr in nlpolter.Tables[0].Rows)
                    {

                        if (float.Parse(dr["Rest_Kubatur"].ToString()) != 0 && dr["FIP"].ToString() == "0")
                        {
                            foreach (PolterResp p in pr.polters)
                            {
                                if (p.foreignRef == dr["ID"].ToString())
                                {
                                    dr["FIP"] = 1;
                                    if (float.Parse(dr["Rest_Kubatur"].ToString()) != float.Parse(p.size) && float.Parse(dr["Kubatur"].ToString()) != float.Parse(p.size))
                                    {
                                        recfromfirma = "";

                                        if (p.recFrom != null)
                                        {
                                            recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                            //recfromfirma wird null gesetzt wenn TryGetValue keinen Fehler findet
                                            dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                        }
                                        else
                                        {
                                            recfromuser = "Unknown";
                                            recfromfirma = "Unknown";
                                        }

                                        //Firma ist bekannt
                                        if (recfromfirma != "Unknown" & recfromfirma != null)
                                        {
                                            sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                                " select " + p.foreignRef + " , 0, '" + p.size + "', getdate(), 723, newid() ");
                                        }

                                        //p.recFrom = null -> Zürichholz hat Polter erstellt
                                        if (recfromfirma == "Unknown")
                                        {
                                            sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                                " select " + p.foreignRef + ", 0, '" + p.size + "', getdate(), 280, newid() ");
                                        }

                                    }
                                }
                            }

                        }
                    }



                    foreach (DataRow dr in nlpolter.Tables[0].Rows)
                    {
                        if (dr["FIP"].ToString() == "0" && dr["Rest_Kubatur"].ToString() == "99999.00")
                        {
                            sqlstr.Append(" insert into AUFTRAG_TRANSPORT(ID_POLTER, KUBATUR, REST_KUBATUR, ANGELEGT, ID_WEBFIRMA_ERRSTELLER, GUID) " +
                                      " select " + dr["ID"].ToString() + ", 0, 0, getdate(), 280, newid() ");
                        }
                    }

                if (af == 0)
                {
                    if (sqlstr.ToString().Length > 0)
                    {
                        sc.SelectSQL(sqlstr.ToString(), "sens");
                    }

                }
                sc.Close();
                client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                return true;
            }
            else
            {
                return false;
            }
        }

        // Polter aus Polver annehmen
        public Boolean accPolter_code()
        {
            //NL
            ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
            sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);

            //Variablen
            DataSet ds1 = new DataSet();
            string username = Settings.Default.uebergabewebfirma;

            //Polver Session
            string text = "Accept PX Polter " + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            
            text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);
            Session s = new Session();
            XPPRDClient client = new XPPRDClient();
            client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
            CreateSessionRequest csr = new CreateSessionRequest();


            csr.application = 2101;
            csr.clientName = Settings.Default.polverusername;
            csr.interfVersion = Settings.Default.polverinterfaceversion;
            CreateSessionResponse csrp = client.CreateSession(csr);
            s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
            s.sessionID = csrp.sessionID;
            text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            //Polver Polter Request
            GetXPViewRequest req = new GetXPViewRequest(s, 2); //Angebotene Polter annehmen; viewType = 2
            GetXPViewResponse pr = client.GetXPView(req);
            text = "AccPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);
            text = "Import Polter in NL " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            foreach (PolterResp p in pr.polters)
            {
                try
                {

                    //Alter Code, Polter werden immer Angenommen
                    //StringBuilder sqlstr = new StringBuilder("if exists(select * from POLTERUEBERGABE where id_polter = (select id from polter where nummer = '" + p.key.polterID + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfrom + "')) and ID_WEBFIRMA_ERRSTELLER = (select ID from WEBFIRMA where FIRMA = LTRIM(RTRIM('" + p.modPrivUser + "'))))" +
                    //                                            " begin select 1 end else" +
                    //                                            " if exists(select * from DISPOSITION where ID_POLTER = (select id from polter where nummer = '" + p.key.polterID + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + p.modPrivUser + "')))" +
                    //                                            " begin select 2 end else" +
                    //                                            " if exists(select * from AUFTRAG_TRANSPORT where id_polter = (select id from polter where nummer = '" + p.key.polterID + "' and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + p.modPrivUser + "')))" +
                    //                                          " begin select 3 end");

                    //ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");

                    //if (ds1 != null)
                    //{
                    //    AccXPBindingRequest acbr = new AccXPBindingRequest(s, p.key, p.ownerRef);
                    //    AccXPBindingResponse acbrr = client.AccXPBinding(acbr);
                    //    text = "Accept Polter Feedback: " + acbrr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                    //    System.IO.File.AppendAllText(Settings.Default.logs, text);
                    //}


                    //Polterfilter
                    if (txtb_EinzelnerPolter.Text != "")
                    {
                        if (p.key.polterID.ToString() == txtb_EinzelnerPolter.Text)
                        {
                            //Annehmen Polter, nur wenn foreignRef nicht gefüllt, dort steht PID drinnen
                            if (p.foreignRef == "")
                            {
                                AccXPBindingRequest acbr = new AccXPBindingRequest(s, p.key, p.ownerRef);
                                AccXPBindingResponse acbrr = client.AccXPBinding(acbr);
                                text = "Accept Polter Feedback: " + acbrr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                System.IO.File.AppendAllText(Settings.Default.logs, text);
                            }
                        }
                    }
                    //Polterfilter ENDE
                    else
                    {
                        //Annehmen Polter, nur wenn foreignRef nicht gefüllt, dort steht PID drinnen
                        if (p.foreignRef == "")
                        {
                            AccXPBindingRequest acbr = new AccXPBindingRequest(s, p.key, p.ownerRef);
                            AccXPBindingResponse acbrr = client.AccXPBinding(acbr);
                            text = "Accept Polter Feedback: " + acbrr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                            System.IO.File.AppendAllText(Settings.Default.logs, text);
                        }
                    }
                }
                catch
                {
                    this.text = "Fehler Accept Polter Polter: PO" + p.key.polterID + " " + DateTime.Now.TimeOfDay + "\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                }
            }



            sc.Close();
            client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
            text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            return true;
        }

        // Polter in Polver ablehnen wenn in Netlogistik Polterübergabe gelöscht wurde
        public Boolean abroPolter_code()
        {
            Boolean cont_ok = getContacts_code(); //Kontakte-Dictionary neu erstellen

            if (cont_ok) //Falls Kontakte OK, dann ist Polterablehnen möglich
            {
                //NL
                ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
                sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);

                //Variablen
                DataSet ds1 = new DataSet();
                String recfromuser, recfromfirma, modprivuser;
                Boolean babro = false;

                this.text = "Start Abbrogate von abgelehnten Poltern " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);

                //Polver Session
                text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                Session s = new Session();
                XPPRDClient client = new XPPRDClient();
                client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
                CreateSessionRequest csr = new CreateSessionRequest();


                csr.application = 2101;
                csr.clientName = Settings.Default.polverusername;
                csr.interfVersion = Settings.Default.polverinterfaceversion;
                CreateSessionResponse csrp = client.CreateSession(csr);
                s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
                s.sessionID = csrp.sessionID;
                text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Polver Polter Request
                GetXPViewRequest req = new GetXPViewRequest(s, 1);
                GetXPViewResponse pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                modprivuser = Settings.Default.webfirma; //Zürichholz Webfirma


                //Polter in Polver ablehnen falls in NL abgelehnt
                foreach (PolterResp p in pr.polters)
                {

                    try
                    {
                        //Polterfilter
                        if (txtb_EinzelnerPolter.Text != "")
                        {
                            if (p.key.polterID.ToString() == txtb_EinzelnerPolter.Text)
                            {
                                recfromfirma = "";
                                babro = false;                             //Boolean, ob Ablehnen von Polter erlaubt ist

                                if (p.recFrom != null)
                                {

                                    recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                    //recfromfirma wird auf null gesetzt falls von TryGetValue nichts gefunden wird
                                    dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                    if (recfromfirma != null)
                                    {
                                        babro = true; //Falls Firma angegeben wurde, aber nicht mittels Dictionary gefunden werden kann, wird er beim Ablehnen übersprungen
                                    }
                                }

                                if (recfromfirma != null & babro) //Abfrage ob recfromfirma bekannt ist
                                {
                                    //Holen von Polter ID und Polterübergabe ID
                                    StringBuilder sqlstr = new StringBuilder("select p.ID as PolterID, pu.ID as PUID from polter p left join polteruebergabe pu on pu.ID_POLTER = p.ID where " +
                                                                          "p.nummer = 'PO" + p.key.polterID + "' AND p.ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')");

                                    ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");

                                    //Wenn Polter ID nicht null ist und Polterübergabe ID null ist, dann ablehnen
                                    if (ds1.Tables[0].Rows[0]["PolterID"].ToString() != null & ds1.Tables[0].Rows[0]["PUID"].ToString() == "") //Wenn keine Polter ID gefunden wurde nicht ablehnen, Wenn Polter ID und Polterübergabe ID gefunden wurden auch nicht
                                    {
                                        AbroXPBindingRequest abbr = new AbroXPBindingRequest(s, p.key);
                                        AbroXPBindingResponse abr = client.AbroXPBinding(abbr);
                                        text = "Abrogate Polter Feedback: " + abr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                        System.IO.File.AppendAllText(Settings.Default.logs, text);
                                    }
                                }

                                if (p.recFrom == null) //recFrom aus Polver ist null -> Zürich hat Polver selbst erstellt
                                {
                                    //Holen von Polter ID und Polterübergabe ID
                                    StringBuilder sqlstr = new StringBuilder("select p.ID as PolterID, pu.ID as PUID from polter p left join polteruebergabe pu on pu.ID_POLTER = p.ID where " +
                                                                          "p.nummer = 'PO" + p.key.polterID + "' AND p.ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')");

                                    ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");

                                    //Wenn Polter ID nicht null ist und Polterübergabe ID null ist, dann ablehnen
                                    if (ds1.Tables[0].Rows[0]["PolterID"].ToString() != null & ds1.Tables[0].Rows[0]["PUID"].ToString() == "") //Wenn keine Polter ID gefunden wurde nicht ablehnen, Wenn Polter ID und Polterübergabe ID gefunden wurden auch nicht
                                    {
                                        AbroXPBindingRequest abbr = new AbroXPBindingRequest(s, p.key);
                                        AbroXPBindingResponse abr = client.AbroXPBinding(abbr);
                                        text = "Abrogate Polter Feedback: " + abr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                        System.IO.File.AppendAllText(Settings.Default.logs, text);
                                    }
                                }
                            }
                        }
                        //Polterfilter ENDE
                        else
                        {
                            recfromfirma = "";
                            babro = false;                             //Boolean, ob Ablehnen von Polter erlaubt ist

                            if (p.recFrom != null)
                            {

                                recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                //recfromfirma wird auf null gesetzt falls von TryGetValue nichts gefunden wird
                                dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                if (recfromfirma != null)
                                {
                                    babro = true; //Falls Firma angegeben wurde, aber nicht mittels Dictionary gefunden werden kann, wird er beim Ablehnen übersprungen
                                }
                            }

                            if (recfromfirma != null & babro) //Abfrage ob recfromfirma bekannt ist
                            {
                                //Holen von Polter ID und Polterübergabe ID
                                StringBuilder sqlstr = new StringBuilder("select p.ID as PolterID, pu.ID as PUID from polter p left join polteruebergabe pu on pu.ID_POLTER = p.ID where " +
                                                                      "p.nummer = 'PO" + p.key.polterID + "' AND p.ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')");

                                ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");

                                //Wenn Polter ID nicht null ist und Polterübergabe ID null ist, dann ablehnen
                                if (ds1.Tables[0].Rows[0]["PolterID"].ToString() != null & ds1.Tables[0].Rows[0]["PUID"].ToString() == "") //Wenn keine Polter ID gefunden wurde nicht ablehnen, Wenn Polter ID und Polterübergabe ID gefunden wurden auch nicht
                                {
                                    AbroXPBindingRequest abbr = new AbroXPBindingRequest(s, p.key);
                                    AbroXPBindingResponse abr = client.AbroXPBinding(abbr);
                                    text = "Abrogate Polter Feedback: " + abr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                    System.IO.File.AppendAllText(Settings.Default.logs, text);
                                }
                            }

                            if (p.recFrom == null) //recFrom aus Polver ist null -> Zürich hat Polver selbst erstellt
                            {
                                //Holen von Polter ID und Polterübergabe ID
                                StringBuilder sqlstr = new StringBuilder("select p.ID as PolterID, pu.ID as PUID from polter p left join polteruebergabe pu on pu.ID_POLTER = p.ID where " +
                                                                      "p.nummer = 'PO" + p.key.polterID + "' AND p.ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + modprivuser + "')");

                                ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");

                                //Wenn Polter ID nicht null ist und Polterübergabe ID null ist, dann ablehnen
                                if (ds1.Tables[0].Rows[0]["PolterID"].ToString() != null & ds1.Tables[0].Rows[0]["PUID"].ToString() == "") //Wenn keine Polter ID gefunden wurde nicht ablehnen, Wenn Polter ID und Polterübergabe ID gefunden wurden auch nicht
                                {
                                    AbroXPBindingRequest abbr = new AbroXPBindingRequest(s, p.key);
                                    AbroXPBindingResponse abr = client.AbroXPBinding(abbr);
                                    text = "Abrogate Polter Feedback: " + abr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                    System.IO.File.AppendAllText(Settings.Default.logs, text);
                                }
                            }

                        }
                        //client.RenewSession(new RenewSessionRequest(s));
                    }
                    catch
                    {
                        this.text = "Fehler Polter abrogate Polter: PO" + p.key.polterID + " " + DateTime.Now.TimeOfDay + "\n";
                        System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                    }
                }


                this.text = "Polter in Polver ablehnen fertiggestellt" + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                sc.Close();
                client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                return true;
            }
            else
            {
                return false;
            }
        }

        // Sortimente von xPolver holen
        public Boolean getAssortments_code()
        {

            ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
            sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);

            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();
            string strSql = "";

            //Laden von HA, HS, GKL von NETLOGISTIK
            try
            {
                dictorgha.Clear();
                dictorghs.Clear();
                dictorggkl.Clear();

                this.text = "Start NL Import der Stammdaten " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                strSql = "select ha.KURZ, ha.ID from WEBFIRMA " +
                    "left join T_HOLZART ha on ha.ID_WEBFIRMA = WEBFIRMA.ID " +
                    "where WEBFIRMA.ID = (select id from webfirma where firma like '" + Settings.Default.webfirma + "') ";
                ds1 = sc.SelectSQL(strSql, "sens");
                ds1.Tables[0].TableName = "Holzart";
                foreach (DataRow dr in ds1.Tables[0].Rows)
                {
                    dictorgha.Add(dr[0].ToString().ToLower(), ((int)dr[1]));
                }
                this.text = "Import der Holzarten erfolgreich " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);

                strSql = "select hs.KURZ, hs.ID from WEBFIRMA " +
                    "left join T_HOLZSORTE hs on hs.ID_WEBFIRMA = WEBFIRMA.ID " +
                    "where WEBFIRMA.ID = (select id from webfirma where firma like '" + Settings.Default.webfirma + "') ";
                ds2 = sc.SelectSQL(strSql, "sens");
                ds2.Tables[0].TableName = "Holzsorte";
                foreach (DataRow dr in ds2.Tables[0].Rows)
                {
                    dictorghs.Add(dr[0].ToString().ToLower(), ((int)dr[1]));
                }
                this.text = "Import der Holzsorten erfolgreich " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);

                strSql = "select gkl.KURZ, gkl.ID from WEBFIRMA " +
                    "left join T_GUETEKLASSE gkl on gkl.ID_WEBFIRMA = WEBFIRMA.ID " +
                    "where WEBFIRMA.ID = (select id from webfirma where firma like '" + Settings.Default.webfirma + "') ";
                ds3 = sc.SelectSQL(strSql, "sens");
                ds3.Tables[0].TableName = "Gueteklasse";
                foreach (DataRow dr in ds3.Tables[0].Rows)
                {
                    dictorggkl.Add(dr[0].ToString().ToLower(), ((int)dr[1]));
                }
                this.text = "Import der Güteklassen erfolgreich " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
            }
            catch
            {
                //MessageBox.Show("Fehler bei Netlogistik-Stammdatenzuordnung!");
                this.text = "Import der NL Stammdaten nicht erfolgreich " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                sc.Close();
                //Abbruch wenn Fehler bei NL Stammdaten
                return false;
            }

            //Erstellen der PX Session
            this.text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, this.text);
            Session s = new Session();
            XPPRDClient client = new XPPRDClient();
            
            client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
            CreateSessionRequest csr = new CreateSessionRequest();

            csr.application = 2101;
            csr.clientName = Settings.Default.polverusername;
            csr.interfVersion = Settings.Default.polverinterfaceversion;
            csr.language = 1;


            CreateSessionResponse csrp = client.CreateSession(csr);
            s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
            s.sessionID = csrp.sessionID;
            this.text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, this.text);

            //Laden der PX Sortimente
            try
            {
                string line;
                int out1, spec1, gkl1, type1;

                this.text = "Import PX Stammdaten " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                GetAssortmentResponse ar = client.GetAssortment(new GetAssortmentRequest(s));
                this.text = "GetAssortments Feedback: " + ar.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);

                dictnlha.Clear();   //Löschen der Dictionarys vor jedem beschreiben
                dictnlhs.Clear();   //
                dictnlgkl.Clear();  //
                dictdm.Clear();     //
                dictlen.Clear();    //

                //Erstellen HA Dictionary
                using (StreamReader sr = new StreamReader(Settings.Default.dictionaryha))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        Int32.TryParse(line.Split('=')[0], out out1);
                        dictnlha.Add(out1, line.Split('=')[1]);
                    }
                    this.text = "Dictionary Holzart erstellt " + DateTime.Now.TimeOfDay + "\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                }
                //Erstellen HS Dictionary
                using (StreamReader sr = new StreamReader(Settings.Default.dictionaryhs))
                {

                    while ((line = sr.ReadLine()) != null)
                    {
                        Int32.TryParse(line.Split('=')[0], out out1);
                        dictnlhs.Add(out1, line.Split('=')[1]);
                    }
                    this.text = "Dictionary Holzsorte erstellt " + DateTime.Now.TimeOfDay + "\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                }
                //Erstellen GKL Dictionary
                using (StreamReader sr = new StreamReader(Settings.Default.dictionarygkl))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        Int32.TryParse(line.Split('=')[0], out out1);
                        dictnlgkl.Add(out1, line.Split('=')[1]);
                    }
                    this.text = "Dictionary Güteklasse erstellt " + DateTime.Now.TimeOfDay + "\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                }

                String dummy1;



                foreach (Assortment a in ar.assortments)
                {
                    System.IO.File.AppendAllText(Settings.Default.logs, a.sortCode + ";" + a.quality + ";" + a.treeSpec + ";" + a.woodType + "\n");


                    //TryGetValue setzt Rückgabewert auf 0 wenn nicht gefunden
                    dictorgha.TryGetValue(a.treeSpec.ToLower(), out spec1);
                    dictorghs.TryGetValue(a.woodType.ToLower(), out type1);
                    dictorggkl.TryGetValue(a.quality.ToLower(), out gkl1);

                    if (spec1 != 0)
                    {   
                        dictnlha.Add(a.sortCode, a.treeSpec);
                    }
                    else
                    {
                        dictnlha.TryGetValue(a.sortCode, out dummy1);

                        if (dummy1 == null)
                        {
                                System.IO.File.AppendAllText(Settings.Default.dictionaryha, "\n" + a.sortCode + "=");
                        }
                        
                    }
                    if (type1 != 0)
                    {
                        dictnlhs.Add(a.sortCode, a.woodType);
                    }
                    else
                    {
                        dictnlhs.TryGetValue(a.sortCode, out dummy1);

                        if (dummy1 == null)
                        {
                            System.IO.File.AppendAllText(Settings.Default.dictionaryhs, "\n" + a.sortCode + "=");
                        }

                    }
                    if (gkl1 != 0)
                    {
                        dictnlgkl.Add(a.sortCode, a.quality);
                    }
                    else
                    {
                        dictnlgkl.TryGetValue(a.sortCode, out dummy1);

                        if (dummy1 == null)
                        {
                            System.IO.File.AppendAllText(Settings.Default.dictionarygkl, "\n" + a.sortCode + "=");
                        }

                    }

                    dictlen.Add(a.sortCode, a.length);
                    dictdm.Add(a.sortCode, a.diameter);

                }

            }
            catch
            {
                this.text = "Fehler PX Stammdaten " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                sc.Close();
                client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //MessageBox.Show("Fehler bei Stammdatenzuordnung!");
                //Abbruch bei Fehler Polverstammdaten
                return false;
            }


            sc.Close();
            client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
            text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);
            //True wenn fehlerfrei durchgelaufen
            return true;
        }

        // Polter Attachments wie Bilder, etc von xPolver holen
        public Boolean getAttachments_code()
        {
            Boolean cont_ok = getContacts_code(); //Kontakt-Dictionary neu aufbauen und abprüfen ob OK
            String recfromfirma, recfromuser;
            int cnt = 0;
            byte[] file;
            FileStream stream;
            BinaryReader reader;

            progressBar1.Value = 80;

            if (cont_ok)
            {
                //NL
                ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
                sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);


                //Polver Session
                text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                Session s = new Session();
                XPPRDClient client = new XPPRDClient();
                client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
                CreateSessionRequest csr = new CreateSessionRequest();

                csr.application = 2101;
                csr.clientName = Settings.Default.polverusername;
                csr.interfVersion = Settings.Default.polverinterfaceversion;
                CreateSessionResponse csrp = client.CreateSession(csr);
                s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
                s.sessionID = csrp.sessionID;
                text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Polver Polter Request
                GetXPViewRequest req = new GetXPViewRequest(s, 1);
                GetXPViewResponse pr = client.GetXPView(req);
                text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);
                text = "Import PolterAttachments in NL " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);

                //Ordner anlegen falls es ihn nicht gibt
                System.IO.Directory.CreateDirectory(Settings.Default.pathtmpdocs + "\\zip");


                //Alle Polter durchgehen mit TryCatch in Schleife für individuellen Polter
                foreach (PolterResp p in pr.polters)
                {

                    //GetXPAttachRequest gar = new GetXPAttachRequest(s, 180);
                    //GetXPAttachResponse garr = client.GetXPAttach(gar);

                    try
                    {
                        //Polterfilter
                        if (txtb_EinzelnerPolter.Text != "")
                        {
                            if (p.key.polterID.ToString() == txtb_EinzelnerPolter.Text)
                            {
                                DetailXPolterRequest dr = new DetailXPolterRequest(s, p.key);
                                DetailXPolterResponse drp = client.DetailXPolter(dr);

                                if (drp.detail.info != null)
                                {
                                    foreach (XPPRD.PolterInfo pi in drp.detail.info)
                                    {
                                        GetXPAttachRequest gar = new GetXPAttachRequest(s, drp.detail.info[cnt].infoID);
                                        GetXPAttachResponse garr = client.GetXPAttach(gar);


                                        if (garr.info != null)
                                        {

                                            if (garr.info.attachment != null)
                                            {
                                                //client.RenewSession(new RenewSessionRequest(s));

                                                System.IO.DirectoryInfo di = new DirectoryInfo(Settings.Default.pathtmpdocs + "\\zip\\");

                                                foreach (FileInfo filed in di.GetFiles())
                                                {
                                                    filed.Delete();
                                                }
                                                foreach (DirectoryInfo dir in di.GetDirectories())
                                                {
                                                    dir.Delete(true);
                                                }

                                                File.WriteAllBytes(Settings.Default.pathtmpdocs + "\\zip\\" + garr.info.attachment.name, garr.info.attachment.content);
                                                File.Delete(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");
                                                ZipFile.CreateFromDirectory(Settings.Default.pathtmpdocs + "\\zip" , Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");
                                               

                                                file = File.ReadAllBytes(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");



                                                StringBuilder hex = new StringBuilder(file.Length * 2);

                                                foreach (byte b in file)
                                                {
                                                    hex.AppendFormat("{0:x2}", b);
                                                }


                                                //Wenn recFrom gefüllt, dann aus Dictionary lesen
                                                if (p.recFrom != null)
                                                {
                                                    recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                                    //Wenn TryGetValue keinen Wert findet wird recfromfirma mit null befüllt
                                                    dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                                }
                                                //wenn recFrom nicht gefüllt, dann auf "Unknown" setzen
                                                else
                                                {
                                                    recfromuser = "Unknown";
                                                    recfromfirma = "Unknown";
                                                }

                                                //Wenn recFrom gefüllt war und eine Adresse gefunden wurde, dass wird der Anhang importiert
                                                if (recfromfirma != "Unknown" & recfromfirma != null)
                                                {
                                                    StringBuilder sqlstr = new StringBuilder("if not exists (select * from polter_files where id_polter = (select id from polter where nummer = 'PO" + p.key.polterID + "'and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')) and uploadfilename = '" + garr.info.attachment.name + ".zip')" +
                                                                                        " begin" +
                                                                                        " insert into polter_files (ID_POLTER, UPLOADFILENAME, UPLOADFILE, UPLOADFILETYP) values" +
                                                                                        " ((select id from polter where nummer = 'PO" + p.key.polterID + "' and ID_WEBFIRMA = (select id from webfirma where firma = '" + recfromfirma + "'))," +
                                                                                        " '" + garr.info.attachment.name + ".zip', 0x" + hex + ", 0)" +
                                                                                        " end");

                                                    sc.SelectSQL(sqlstr.ToString(), "sens");
                                                }

                                                //Wenn recFrom leer ist, ist Polter von Zürichholz
                                                if (recfromfirma == "Unknown")
                                                {
                                                    StringBuilder sqlstr = new StringBuilder("if not exists (select * from polter_files where id_polter = (select id from polter where nummer = 'PO" + p.key.polterID + "'and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "')) and uploadfilename = '" + garr.info.attachment.name + ".zip')" +
                                                                                        " begin" +
                                                                                        " insert into polter_files (ID_POLTER, UPLOADFILENAME, UPLOADFILE, UPLOADFILETYP) values" +
                                                                                        " ((select id from polter where nummer = 'PO" + p.key.polterID + "' and ID_WEBFIRMA = (select id from webfirma where firma = '" + Settings.Default.uebergabewebfirma + "'))," +
                                                                                        " '" + garr.info.attachment.name + ".zip', 0x" + hex + ", 0)" +
                                                                                        " end");

                                                    sc.SelectSQL(sqlstr.ToString(), "sens");
                                                }

                                                File.Delete(Settings.Default.pathtmpdocs + "\\zip\\" + garr.info.attachment.name);
                                                File.Delete(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");

                                            }
                                        }

                                        cnt++;

                                        text = "Get Attachment Feedback: " + garr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                        System.IO.File.AppendAllText(Settings.Default.logs, text);
                                    }
                                }
                                cnt = 0;

                            }
                        }
                        //Polterfilter ENDE
                        else
                        {
                            DetailXPolterRequest dr = new DetailXPolterRequest(s, p.key);
                            DetailXPolterResponse drp = client.DetailXPolter(dr);

                            if (drp.detail.info != null)
                            {
                                foreach (XPPRD.PolterInfo pi in drp.detail.info)
                                {
                                    GetXPAttachRequest gar = new GetXPAttachRequest(s, drp.detail.info[cnt].infoID);
                                    GetXPAttachResponse garr = client.GetXPAttach(gar);


                                    if (garr.info != null)
                                    {

                                        if (garr.info.attachment != null)
                                        {
                                            System.IO.DirectoryInfo di = new DirectoryInfo(Settings.Default.pathtmpdocs + "\\zip\\");

                                            foreach (FileInfo filed in di.GetFiles())
                                            {
                                                filed.Delete();
                                            }
                                            foreach (DirectoryInfo dir in di.GetDirectories())
                                            {
                                                dir.Delete(true);
                                            }

                                            File.WriteAllBytes(Settings.Default.pathtmpdocs + "\\zip\\" + garr.info.attachment.name, garr.info.attachment.content);
                                            File.Delete(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");
                                            ZipFile.CreateFromDirectory(Settings.Default.pathtmpdocs + "\\zip", Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");



                                            file = File.ReadAllBytes(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");

                                            StringBuilder hex = new StringBuilder(file.Length * 2);

                                            foreach (byte b in file)
                                            {
                                                hex.AppendFormat("{0:x2}", b);
                                            }

                                            //Wenn recFrom gefüllt, dann aus Dictionary lesen
                                            if (p.recFrom != null)
                                            {
                                                recfromuser = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                                                //Wenn TryGetValue keinen Wert findet wird recfromfirma mit null befüllt
                                                dictcontacts.TryGetValue(recfromuser, out recfromfirma);
                                            }
                                            //wenn recFrom nicht gefüllt, dann auf "Unknown" setzen
                                            else
                                            {
                                                recfromuser = "Unknown";
                                                recfromfirma = "Unknown";
                                            }

                                            //Wenn recFrom gefüllt war und eine Adresse gefunden wurde, dass wird der Anhang importiert
                                            if (recfromfirma != "Unknown" & recfromfirma != null)
                                            {
                                                StringBuilder sqlstr = new StringBuilder("if not exists (select * from polter_files where id_polter = (select id from polter where nummer = 'PO" + p.key.polterID + "'and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + recfromfirma + "')) and uploadfilename = '" + garr.info.attachment.name + ".zip')" +
                                                                                    " begin" +
                                                                                    " insert into polter_files (ID_POLTER, UPLOADFILENAME, UPLOADFILE, UPLOADFILETYP) values" +
                                                                                    " ((select id from polter where nummer = 'PO" + p.key.polterID + "' and ID_WEBFIRMA = (select id from webfirma where firma = '" + recfromfirma + "'))," +
                                                                                    " '" + garr.info.attachment.name + ".zip', 0x" + hex + ", 0)" +
                                                                                    " end");

                                                sc.SelectSQL(sqlstr.ToString(), "sens");
                                            }

                                            //Wenn recFrom leer ist, ist Polter von Zürichholz
                                            if (recfromfirma == "Unknown")
                                            {
                                                StringBuilder sqlstr = new StringBuilder("if not exists (select * from polter_files where id_polter = (select id from polter where nummer = 'PO" + p.key.polterID + "'and ID_WEBFIRMA = (select ID from WEBFIRMA where FIRMA = '" + Settings.Default.uebergabewebfirma + "')) and uploadfilename = '" + garr.info.attachment.name + ".zip')" +
                                                                                    " begin" +
                                                                                    " insert into polter_files (ID_POLTER, UPLOADFILENAME, UPLOADFILE, UPLOADFILETYP) values" +
                                                                                    " ((select id from polter where nummer = 'PO" + p.key.polterID + "' and ID_WEBFIRMA = (select id from webfirma where firma = '" + Settings.Default.uebergabewebfirma + "'))," +
                                                                                    " '" + garr.info.attachment.name + ".zip', 0x" + hex + ", 0)" +
                                                                                    " end");

                                                sc.SelectSQL(sqlstr.ToString(), "sens");
                                            }

                                            File.Delete(Settings.Default.pathtmpdocs + "\\zip\\" + garr.info.attachment.name);
                                            File.Delete(Settings.Default.pathtmpdocs + "\\" + garr.info.attachment.name + ".zip");

                                        }
                                    }

                                    cnt++;

                                    text = "Get Attachment Feedback: " + garr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
                                    System.IO.File.AppendAllText(Settings.Default.logs, text);
                                }
                            }
                            cnt = 0;

                        }
                    }
                    catch
                    {
                        this.text = "Fehler Attachments bei Polter PO:" + p.key.polterID + " " + DateTime.Now.TimeOfDay + "\n";
                        System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                    }

                    //client.RenewSession(new RenewSessionRequest(s));

                }

                sc.Close();
                client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                System.IO.File.AppendAllText(Settings.Default.logs, text);


                progressBar1.Value = 100;

                return true;
            }
            else
            {
                progressBar1.Value = 100;
                return false;
            }

        }

        // Kontakte von Dictionary und xPolver holen
        public Boolean getContacts_code()
        {
            String line, contactnl, recfromfirma;

            //NL
            ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
            sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);


            //Polver Session
            text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);
            Session s = new Session();
            XPPRDClient client = new XPPRDClient();

            client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
            CreateSessionRequest csr = new CreateSessionRequest();

            csr.application = 2101;
            csr.clientName = Settings.Default.polverusername;
            csr.interfVersion = Settings.Default.polverinterfaceversion;
            CreateSessionResponse csrp = client.CreateSession(csr);
            s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
            s.sessionID = csrp.sessionID;
            text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            //Polver Polter Request
            GetXPViewRequest req = new GetXPViewRequest(s, 1);
            GetXPViewResponse pr = client.GetXPView(req);
            text = "GetPolter Feedback: " + pr.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            //Kontakte löschen um neu anzufangen
            dictcontacts.Clear();

            

            //Erstellen Kontakt Dictionary
            using (StreamReader sr = new StreamReader(Settings.Default.dictionarycontacts))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    dictcontacts.Add(line.Split('=')[0], line.Split('=')[1]);
                }
                this.text = "Dictionary Kontakte erstellt " + DateTime.Now.TimeOfDay + "\n";
                System.IO.File.AppendAllText(Settings.Default.logs, this.text);
            }

            //RecFrom von Poltern schreiben
            foreach (PolterResp p in pr.polters)
            {

                try
                {
                   
                    if (p.recFrom != null)
                    {
                        //Klammern, Nummern und leerzeichen am Schluss von recFrom entfernen
                        recfromfirma = Regex.Replace(p.recFrom, "[()0-9]", "").Trim();
                        dictcontacts.TryGetValue(recfromfirma, out contactnl);

                        try
                        {
                            //Wenn recFrom noch nicht in NL Dictionary ist, dann Dictionary löschen, neuen Wert hinzufügen und Werte aus File anfügen
                            if (dictcontacts.Keys.Contains(recfromfirma) == false)
                            {
                                dictcontacts.Clear();
                                System.IO.File.AppendAllText(Settings.Default.dictionarycontacts, recfromfirma + "=" + "\n");

                                //EMAIL
                                MailMessage mail = new MailMessage(Settings.Default.mail_from, Settings.Default.mail_to);
                                SmtpClient smtpclient = new SmtpClient();
                                //System.Net.NetworkCredential NC = new System.Net.NetworkCredential("qms@winforstpro.com", "netlog123");
                                System.Net.NetworkCredential NC = new System.Net.NetworkCredential(Settings.Default.smtp_user, Settings.Default.smtp_pw);

                                smtpclient.UseDefaultCredentials = false;
                                smtpclient.Port = int.Parse(Settings.Default.smtp_port);
                                smtpclient.Host = Settings.Default.smtp_host;
                                //smtpclient.Host = "smtp.1und1.de";
                                smtpclient.Credentials = NC;
                                mail.Subject = "xPolver Schnittstelle Adresszuordnung";
                                mail.Body = "Adressen Dictionary prüfen!";
                                smtpclient.Send(mail);

                                using (StreamReader sr = new StreamReader(Settings.Default.dictionarycontacts))
                                {

                                    //System.IO.File.AppendAllText(Settings.Default.dictionarycontacts, repl + "=" + "\n");
                                    while ((line = sr.ReadLine()) != null)
                                    {
                                        dictcontacts.Add(line.Split('=')[0], line.Split('=')[1]);
                                    }
                                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);

                                }

                            }
                        }

                        catch
                        {
                            //MessageBox.Show("Es ist ein Problem beim Email Versand aufgetreten!");
                        }

                        
                    }
                    //client.RenewSession(new RenewSessionRequest(s));
                }
                catch
                {
                    sc.Close();
                    client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
                    text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, text);
                    //Bei Fehler wird ganz abgebrochen
                    //MessageBox.Show("Adresszuordnung nicht erfolgreich abgeschlossen!");
                    return false;
                }
            }

            sc.Close();
            client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
            text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            //True wenn fehlerfrei durchgelaufen
            return true;
        }

        // Programm einmal komplett inclusive Stammdaten durchlaufen lassen
        public void btnStart_Click(object sender, EventArgs e)
        {

            //getAssortments_code();
            //getContacts_code();
            //accPolter_code();
            getPolter_code(); //getAssortments, getContacts und accPolter werden im getPolter ausgeführt
            getAttachments_code(); //
            progressBar1.Value = 90;
            setPolter_code(); //
            progressBar1.Value = 95;
            abroPolter_code(); //

            progressBar1.Value = 100;
        }

        // Create Hash value out of Password and Salt
        public static string createSHA256(string sPassword_p, string sSalt_p)
        {
            string sStr = sPassword_p + sSalt_p;
            SHA256Managed crypt = new SHA256Managed();
            string sHash = "";
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(sStr), 0, Encoding.UTF8.GetByteCount(sStr));
            for (int i = 0; i < SHA256_DIGEST_LENGTH; i++)
            {
                sHash += String.Format("{0:x2}", crypto[i]);
            }
            return sHash + sSalt_p;
        }

        // Speichern der Einstellungen
        private void btnSave_Click(object sender, EventArgs e)
        {
            Settings.Default.dictionaryha = txtb_dictha.Text;
            Settings.Default.dictionaryhs = txtb_dicths.Text;
            Settings.Default.dictionarygkl = txtb_dictgkl.Text;
            Settings.Default.webservice = txtb_service.Text;
            Settings.Default.client_endpoint = txtb_clientendpoint.Text;
            Settings.Default.uebergabewebfirma = txtb_username.Text;
            Settings.Default.webfirma = txtb_password.Text;
            Settings.Default.polverinterfaceversion = txtb_pxinterfversion.Text;
            Settings.Default.polverusername = txtb_pxusername.Text;
            Settings.Default.polverpassword = txtb_pxpw.Text;
            Settings.Default.logs = txtb_logfile.Text;
            Settings.Default.dictionarycontacts = txtb_contacts.Text;
            Settings.Default.mail_from = txtb_mailfrom.Text;
            Settings.Default.mail_to = txtb_mailto.Text;
            Settings.Default.smtp_host = txtb_smtphost.Text;
            Settings.Default.smtp_port = txtb_smtpport.Text;
            Settings.Default.smtp_user = txtb_smtpuser.Text;
            Settings.Default.smtp_pw = txtb_smtppw.Text;
            Settings.Default.pathtmpdocs = txtb_tmpdocs.Text;
            Settings.Default.Save();
        }


        //Rücksynchronisation an Polver
        private void btnSendPolter_Click(object sender, EventArgs e)
        {

            progressBar1.Value = 0;
            //NL
            ConsoleApplication1.WfpNetService.ServiceClient sc = new ConsoleApplication1.WfpNetService.ServiceClient();
            sc.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.webservice);


            //Variablen
            DataSet ds1 = new DataSet();
            string sqlstr;
            int i = 0, cosemat = 1;

            //session
            text = "Create PX Session " + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);
            Session s = new Session();
            XPPRDClient client = new XPPRDClient();
            client.Endpoint.Address = new System.ServiceModel.EndpointAddress(Settings.Default.client_endpoint);
            CreateSessionRequest csr = new CreateSessionRequest();

            csr.application = 2101;
            csr.clientName = Settings.Default.polverusername;
            csr.interfVersion = Settings.Default.polverinterfaceversion;
            CreateSessionResponse csrp = client.CreateSession(csr);
            s.hash = createSHA256(Settings.Default.polverpassword, csrp.salt.ToString());
            s.sessionID = csrp.sessionID;
            text = "Session Feedback: " + csrp.fb.text + ' ' + DateTime.Now.TimeOfDay + "\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            XPPRD.PolterReq pr = new XPPRD.PolterReq();
            XPPRD.Point xpoint = new XPPRD.Point();

            sqlstr = new StringBuilder("select POLTER.ID as PID, isnull(KOORDINATE_X,0) as KOORDINATE_X, isnull(KOORDINATE_Y,0) as KOORDINATE_Y, isnull(POLTER.BEMERKUNG, '') as BEMERKUNG, isnull(THA.KURZ, '') as HA, isnull(THS.KURZ, '') as HS, isnull(TGK.KURZ, '') as GKL, isnull(TSK.KURZ, '') as STKL" +
                                        " , isnull(PP.LANGE, 0) as LANGE, isnull(PP.STUCK, 0) as STUCK, isnull(PP.KUBATUR, 0) as KUBATUR, WEBFIRMA.FIRMA, isnull(PP.COSEMAT, 0) as COSEMAT" +
                                        " , isnull(WEBFIRMA.POLVERUEBERGABE, 0) as POLVERUEBERGABE, isnull(ADR.KURZBEZEICHNUNG, '') as SAEGEWERK, isnull(ADR_WB.KURZBEZEICHNUNG, '') AS WALDBESITZER, polter.NUMMER as POLTERNR, Polter.LOS" +
                                        " from POLTER" +
                                        " left join POLTERPOOL PP on PP.ID_POLTER = POLTER.ID" +
                                        " left join POLTERUEBERGABE PU on PU.ID_POLTER = POLTER.ID" +
                                        " left join T_HOLZART THA on THA.ID = PP.ID_T_HOLZART" +
                                        " left join T_HOLZSORTE THS on THS.ID = PP.ID_T_HOLZSORTE" +
                                        " left join T_GUETEKLASSE TGK on TGK.ID = PP.ID_T_GUETEKLASSE" +
                                        " left join T_STKL TSK on TSK.ID = PP.ID_T_STKL" +
                                        " left join WEBFIRMA on WEBFIRMA.ID = PU.ID_WEBFIRMA" +
                                        " left join ADR on ADR.ID = PP.ID_ADR_SW" +
                                        " LEFT JOIN ADR ADR_WB ON ADR_WB.ID = PP.ID_ADR_WALD" +
                                        " where PU.ID_WEBFIRMA_ERRSTELLER = 280 and webfirma.polveruebergabe = 1").ToString();
            ds1 = sc.SelectSQL(sqlstr.ToString(), "sens");


            progressBar1.Value = 50;

            while (i < ds1.Tables[0].Rows.Count)
            {

                try
                {
                    cosemat = Int32.Parse(ds1.Tables[0].Rows[i]["COSEMAT"].ToString());

                    if (cosemat != 1)
                    {
                        //Falls Koordinaten im WGS84 Format sind
                        if(double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_X"].ToString()) >= 180)
                        {
                            xpoint.posE = (CHtoWGSlng(Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_X"].ToString()), Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_Y"].ToString())) * 1000000).ToString(); 
                            xpoint.posN = (CHtoWGSlat(Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_X"].ToString()), Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_Y"].ToString())) * 1000000).ToString();

                        }
                        else
                        {
                            xpoint.posE = (Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_X"].ToString())* 1000000).ToString();
                            xpoint.posN = (Double.Parse(ds1.Tables[0].Rows[i]["KOORDINATE_Y"].ToString()) * 1000000).ToString();
                        }

                        xpoint.projection = 2;

                        pr.size = (int)float.Parse(ds1.Tables[0].Rows[i]["KUBATUR"].ToString());
                        pr.sortCode = 499;
                        pr.polterInfo = ds1.Tables[0].Rows[i]["HS"].ToString() + " " + ds1.Tables[0].Rows[i]["HA"].ToString() + " " + ds1.Tables[0].Rows[i]["GKL"].ToString() + " " + ds1.Tables[0].Rows[i]["STKL"].ToString();
                        if (float.Parse(ds1.Tables[0].Rows[i]["LANGE"].ToString()) != 0)
                        {
                            pr.polterInfo = pr.polterInfo + " " + ds1.Tables[0].Rows[i]["LANGE"].ToString();
                        }
                        if (float.Parse(ds1.Tables[0].Rows[i]["STUCK"].ToString()) != 0)
                        {
                            pr.polterInfo = pr.polterInfo + " " + ds1.Tables[0].Rows[i]["STUCK"].ToString();
                        }
                        if (ds1.Tables[0].Rows[i]["SAEGEWERK"].ToString() != "")
                        {
                            pr.polterInfo = pr.polterInfo + " " + ds1.Tables[0].Rows[i]["SAEGEWERK"].ToString();
                        }
                        if (ds1.Tables[0].Rows[i]["BEMERKUNG"].ToString() != "")
                        {
                            pr.polterInfo = pr.polterInfo + " " + ds1.Tables[0].Rows[i]["BEMERKUNG"].ToString();
                        }
                        

                        pr.foreignRef = ds1.Tables[0].Rows[i]["PID"].ToString();
                        pr.sizeSpecified = true;
                        pr.sortCodeSpecified = true;
                        pr.addInfo1 = ds1.Tables[0].Rows[i]["POLTERNR"].ToString(); //Polternummer
                        pr.addInfo2 = ds1.Tables[0].Rows[i]["LOS"].ToString(); //Losnr
                        pr.ownerInfo = ds1.Tables[0].Rows[i]["WALDBESITZER"].ToString(); //Waldbesitzer

                        CreateXPolterRequest req = new CreateXPolterRequest(s, xpoint, pr);
                        CreateXPolterResponse res = client.CreateXPolter(req);

                        if (res.polter != null)
                        {
                            sqlstr = new StringBuilder("update polterpool set cosemat = 1 where id_polter = " + ds1.Tables[0].Rows[i]["PID"].ToString()).ToString();
                            sc.SelectSQL(sqlstr.ToString(), "sens");

                            DataSet dsPolterfiles = new DataSet();
                            PolterInfo pi = new PolterInfo();
                            AddXPAttachRequest areq;
                            AddXPAttachResponse arp;
                            Byte[] ulfile;


                            sqlstr = new StringBuilder("SELECT * from polter_files where ID_POLTER = " + ds1.Tables[0].Rows[i]["PID"].ToString()).ToString();
                            dsPolterfiles = sc.SelectSQL(sqlstr.ToString(), "sens");

                            foreach (DataRow row in dsPolterfiles.Tables[0].Rows)
                            {
                                pi.infoCode = 1004;
                                pi.infoIDSpecified = true;
                                pi.infoText = row["uploadfilename"].ToString();

                                XPPRD.Binary polterfile = new XPPRD.Binary();
                                
                                ulfile = (byte[])row["uploadfile"];

                                File.WriteAllBytes(Settings.Default.pathtmpdocs + "\\zip\\" + row["uploadfilename"].ToString(),ulfile);


                                polterfile.content = ulfile;
                                polterfile.name = row["uploadfilename"].ToString();
                                pi.attachment = polterfile;
                                pi.attachment.mimeType = "application/zip";
                                areq = new AddXPAttachRequest(s, res.polter.key, 0, pi);
                                arp = client.AddXPAttach(areq);
                            }

                        }
                        else
                        {
                            //MessageBox.Show("Polter aus PID: " + ds1.Tables[0].Rows[i]["PID"].ToString() + " konnte nicht im Polver erstellt werden. " + res.fb.text);
                            this.text = "Polter aus PID: " + ds1.Tables[0].Rows[i]["PID"].ToString() + " konnte nicht im Polver erstellt werden. " + res.fb.text + " " + DateTime.Now.TimeOfDay + "\n";
                            System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                        }

                        

                    }
                    i++;
                }
                catch
                {
                    this.text = "Fehler Poltererstellung in Polver: PID" + ds1.Tables[0].Rows[i]["PID"].ToString() + " " + DateTime.Now.TimeOfDay + "\n";
                    System.IO.File.AppendAllText(Settings.Default.logs, this.text);
                    i++;
                }
            }


            sc.Close();
            client.CloseSession(new CloseSessionRequest(s.sessionID, csrp.salt));
            text = "Close Session: " + DateTime.Now.TimeOfDay + "\n\n";
            System.IO.File.AppendAllText(Settings.Default.logs, text);

            progressBar1.Value = 100;
        }

        public static byte[] ObjectToByteArray(Object obj)
        {
            BinaryFormatter bf = new BinaryFormatter();
            using (var ms = new MemoryStream())
            {
                bf.Serialize(ms, obj);
                return ms.ToArray();
            }
        }

        public void accPolter(object sender, EventArgs e)
        {
            accPolter_code();
        }
        public void getContacts(object sender, EventArgs e)
        {
            getContacts_code();
        }
        public void getAttachments(object sender, EventArgs e)
        {
            getAttachments_code();
        }
        public void getAssortments(object sender, EventArgs e)
        {
            getAssortments_code();
        }
        public void abroPolter(object sender, EventArgs e)
        {
            abroPolter_code();
        }
        public void setPolter(object sender, EventArgs e)
        {
            setPolter_code();
        }
        public void getPolter(object sender, EventArgs e)
        {
            getPolter_code();
        }

        // Convert angle in decimal degrees to sexagesimal seconds
        public static double DecToSexAngle(double dec)
        {
            int deg = (int)Math.Floor(dec);
            int min = (int)Math.Floor((dec - deg) * 60.0);
            double sec = (((dec - deg) * 60.0) - min) * 60.0;

            return sec + (double)min * 60.0 + (double)deg * 3600.0;
        }

        // Convert WGS lat/long (° dec) to CH y
        private static double WGStoCHy(double lat, double lng)
        {
            // Convert decimal degrees to sexagesimal seconds
            lat = DecToSexAngle(lat);
            lng = DecToSexAngle(lng);

            // Auxiliary values (% Bern)
            double lat_aux = (lat - 169028.66) / 10000.0;
            double lng_aux = (lng - 26782.5) / 10000.0;

            // Process Y
            double y = 600072.37
                + 211455.93 * lng_aux
                - 10938.51 * lng_aux * lat_aux
                - 0.36 * lng_aux * Math.Pow(lat_aux, 2)
                - 44.54 * Math.Pow(lng_aux, 3);

            return y;
        }

        // Convert WGS lat/long (° dec) to CH x
        private static double WGStoCHx(double lat, double lng)
        {
            // Convert decimal degrees to sexagesimal seconds
            lat = DecToSexAngle(lat);
            lng = DecToSexAngle(lng);

            // Auxiliary values (% Bern)
            double lat_aux = (lat - 169028.66) / 10000.0;
            double lng_aux = (lng - 26782.5) / 10000.0;

            // Process X
            double x = 200147.07
                + 308807.95 * lat_aux
                + 3745.25 * Math.Pow(lng_aux, 2)
                + 76.63 * Math.Pow(lat_aux, 2)
                - 194.56 * Math.Pow(lng_aux, 2) * lat_aux
                + 119.79 * Math.Pow(lat_aux, 3);

            return x;
        }

        // Convert CH y/x to WGS lat
        private static double CHtoWGSlat(double y, double x)
        {
            // Converts military to civil and  to unit = 1000km
            // Auxiliary values (% Bern)
            double y_aux = (y - 600000.0) / 1000000.0;
            double x_aux = (x - 200000.0) / 1000000.0;

            // Process lat
            double lat = 16.9023892
                + 3.238272 * x_aux
                - 0.270978 * Math.Pow(y_aux, 2)
                - 0.002528 * Math.Pow(x_aux, 2)
                - 0.0447 * Math.Pow(y_aux, 2) * x_aux
                - 0.0140 * Math.Pow(x_aux, 3);

            // Unit 10000" to 1 " and converts seconds to degrees (dec)
            lat = lat * 100 / 36;

            return lat;
        }

        // Convert CH y/x to WGS long
        private static double CHtoWGSlng(double y, double x)
        {
            // Converts military to civil and  to unit = 1000km
            // Auxiliary values (% Bern)
            double y_aux = (y - 600000.0) / 1000000.0;
            double x_aux = (x - 200000.0) / 1000000.0;

            // Process long
            double lng = 2.6779094
                + 4.728982 * y_aux
                + 0.791484 * y_aux * x_aux
                + 0.1306 * y_aux * Math.Pow(x_aux, 2)
                - 0.0436 * Math.Pow(y_aux, 3);

            // Unit 10000" to 1 " and converts seconds to degrees (dec)
            lng = lng * 100.0 / 36.0;

            return lng;
        }

    }


}
