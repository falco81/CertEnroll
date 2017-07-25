using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//  Add the CertEnroll namespace
using CERTENROLLLib;
using CERTCLIENTLib;
using System.Xml;
using System.Security.Cryptography.X509Certificates;
using System.IO;



namespace CertEnroll
{
    public partial class Form1 : Form
    {
        private const int CC_DEFAULTCONFIG = 0;
        private const int CC_UIPICKCONFIG = 0x1;
        private const int CR_IN_BASE64 = 0x1;
        private const int CR_IN_FORMATANY = 0;
        private const int CR_IN_PKCS10 = 0x100;
        private const int CR_DISP_ISSUED = 0x3;
        private const int CR_DISP_UNDER_SUBMISSION = 0x5;
        private const int CR_OUT_BASE64 = 0x1;
        private const int CR_OUT_CHAIN = 0x100;
        string strRequest;
        string strCert;
        string myTemplate;
        string CAName;
        string ExportPass;
        Byte[] certData;

        public Form1()
        {
            InitializeComponent();
        }

        // Create request
        private void createRequestButton_Click(object sender, EventArgs e)
        {
            createRequestButton.Enabled = false;
            exportButton.Enabled = false;
            label3.Text = "Pracuji...";
            //  Create all the objects that will be required
            CX509CertificateRequestPkcs10 objPkcs10 = new CX509CertificateRequestPkcs10();
            CX509PrivateKey objPrivateKey = new CX509PrivateKeyClass();
            CCspInformation objCSP = new CCspInformationClass();
            CCspInformations objCSPs = new CCspInformationsClass();
            CX500DistinguishedName objDN = new CX500DistinguishedNameClass();
            CX509Enrollment objEnroll = new CX509EnrollmentClass();
            CObjectIds objObjectIds = new CObjectIdsClass();
            CObjectId objObjectId = new CObjectIdClass();
            CX509ExtensionKeyUsage objExtensionKeyUsage = new CX509ExtensionKeyUsageClass();
            CX509ExtensionEnhancedKeyUsage objX509ExtensionEnhancedKeyUsage = new CX509ExtensionEnhancedKeyUsageClass();
            

            try
            {
//                requestText.Text = "";
                strRequest = "";

                //  Initialize the csp object using the desired Cryptograhic Service Provider (CSP)
                objCSP.InitializeFromName(
                    "Microsoft Enhanced Cryptographic Provider v1.0"
                );

                //  Add this CSP object to the CSP collection object
                objCSPs.Add(
                    objCSP
                );

                //  Provide key container name, key length and key spec to the private key object
                //objPrivateKey.ContainerName = "AlejaCMa";
                objPrivateKey.Length = 2048;
                objPrivateKey.KeySpec = X509KeySpec.XCN_AT_SIGNATURE;
                objPrivateKey.KeyUsage = X509PrivateKeyUsageFlags.XCN_NCRYPT_ALLOW_ALL_USAGES;
                objPrivateKey.ExportPolicy = X509PrivateKeyExportFlags.XCN_NCRYPT_ALLOW_EXPORT_FLAG;
                objPrivateKey.MachineContext = true;

                //  Provide the CSP collection object (in this case containing only 1 CSP object)
                //  to the private key object
                objPrivateKey.CspInformations = objCSPs;

                //  Create the actual key pair
                objPrivateKey.Create();

                //  Initialize the PKCS#10 certificate request object based on the private key.
                //  Using the context, indicate that this is a user certificate request and don't
                //  provide a template name
                objPkcs10.InitializeFromPrivateKey(
                    X509CertificateEnrollmentContext.ContextMachine,
                    objPrivateKey,
                    myTemplate
                );

                // Key Usage Extension 
/*                objExtensionKeyUsage.InitializeEncode(
                    X509KeyUsageFlags.XCN_CERT_DIGITAL_SIGNATURE_KEY_USAGE |
                    X509KeyUsageFlags.XCN_CERT_NON_REPUDIATION_KEY_USAGE |
                    X509KeyUsageFlags.XCN_CERT_KEY_ENCIPHERMENT_KEY_USAGE |
                    X509KeyUsageFlags.XCN_CERT_DATA_ENCIPHERMENT_KEY_USAGE
                );
                objPkcs10.X509Extensions.Add((CX509Extension)objExtensionKeyUsage);

                // Enhanced Key Usage Extension
                objObjectId.InitializeFromValue("1.3.6.1.5.5.7.3.2"); // OID for Client Authentication usage
                objObjectIds.Add(objObjectId);
                objX509ExtensionEnhancedKeyUsage.InitializeEncode(objObjectIds);
                objPkcs10.X509Extensions.Add((CX509Extension)objX509ExtensionEnhancedKeyUsage);
*/                

                //  Encode the name in using the Distinguished Name object
                objDN.Encode(
                    "CN="+textBox1.Text+",OU="+comboBox1.Text,
                    X500NameFlags.XCN_CERT_NAME_STR_NONE
                );

                //  Assing the subject name by using the Distinguished Name object initialized above
                objPkcs10.Subject = objDN;

                CAlternativeName objRfc822Name = new CAlternativeName(); 
                CAlternativeNames objAlternativeNames = new CAlternativeNames(); 
                CX509ExtensionAlternativeNames objExtensionAlternativeNames = new CX509ExtensionAlternativeNames(); 
                // Set Alternative RFC822 Name 
                objRfc822Name.InitializeFromString(AlternativeNameType.XCN_CERT_ALT_NAME_DNS_NAME, textBox1.Text); 
                // Set Alternative Names 
                objAlternativeNames.Add(objRfc822Name); 
                objExtensionAlternativeNames.InitializeEncode(objAlternativeNames); 
                objPkcs10.X509Extensions.Add((CX509Extension)objExtensionAlternativeNames);

                // Create enrollment request
                objEnroll.InitializeFromRequest(objPkcs10);              
                strRequest = objEnroll.CreateRequest(
                    EncodingType.XCN_CRYPT_STRING_BASE64
                );

//                requestText.Text = strRequest;
                //MessageBox.Show("Žádost vygenerovaná");
                label3.Text = "Žádost vygenerovaná";
                sendRequestButton.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                createRequestButton.Enabled = true;
                sendRequestButton.Enabled = false;
                acceptPKCS7Button.Enabled = false;
            }

        }


        // Submit request to CA and get response 
        private void sendRequestButton_Click(object sender, EventArgs e)
        {
            sendRequestButton.Enabled = false;
            label3.Text = "Pracuji...";
            //  Create all the objects that will be required
            CCertConfig objCertConfig = new CCertConfigClass();
            CCertRequest objCertRequest = new CCertRequestClass();
            string strCAConfig;
//            string strRequest;
            int iDisposition;
            string strDisposition;


            try
            {
//                strRequest = requestText.Text;

                // Get CA config from UI
                //strCAConfig = objCertConfig.GetConfig(CC_DEFAULTCONFIG);
                strCAConfig = objCertConfig.GetConfig(CC_UIPICKCONFIG);

                // Submit the request
                iDisposition = objCertRequest.Submit(
                    CR_IN_BASE64 | CR_IN_FORMATANY,
                    strRequest,
                    null,
                    strCAConfig
                );

                // Check the submission status
                if (CR_DISP_ISSUED != iDisposition) // Not enrolled
                {
                    strDisposition = objCertRequest.GetDispositionMessage();

                    if (CR_DISP_UNDER_SUBMISSION == iDisposition) // Pending
                    {
                        MessageBox.Show("The submission is pending: " + strDisposition);
                        return;
                    }
                    else // Failed
                    {
                        MessageBox.Show("The submission failed: " + strDisposition);
                        MessageBox.Show("Last status: " + objCertRequest.GetLastStatus().ToString());
                        return;
                    }
                }

                // Get the certificate
                strCert = objCertRequest.GetCertificate(
                    CR_OUT_BASE64 | CR_OUT_CHAIN
                );
                  //MessageBox.Show("Žádost odeslaná");
                  label3.Text = "Žádost odeslaná";
                  acceptPKCS7Button.Enabled = true;
 //               responseText.Text = strCert;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                createRequestButton.Enabled = false;
                sendRequestButton.Enabled = true;
                acceptPKCS7Button.Enabled = false;

            }

        }


        // Install response from CA
        private void acceptPKCS7Button_Click(object sender, EventArgs e)
        {
            acceptPKCS7Button.Enabled = false;
            label3.Text = "Pracuji...";
            //  Create all the objects that will be required
            CX509Enrollment objEnroll = new CX509EnrollmentClass();
//            string strCert;

            try
            {
//                strCert = responseText.Text;

                // Install the certificate
                objEnroll.Initialize(X509CertificateEnrollmentContext.ContextMachine);
                objEnroll.InstallResponse(
                    InstallResponseRestrictionFlags.AllowUntrustedRoot,
                    strCert,
                    EncodingType.XCN_CRYPT_STRING_BASE64,
                    null
                );

                //MessageBox.Show("Certificate installed!");
                label3.Text = "Certifikát nainstalován do machine uložiště";
                createRequestButton.Enabled = true;
                exportButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                createRequestButton.Enabled = false;
                sendRequestButton.Enabled = false;
                acceptPKCS7Button.Enabled = true;
            }

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            sendRequestButton.Enabled = false;
            acceptPKCS7Button.Enabled = false;
            exportButton.Enabled = false;
            //comboBox1.SelectedItem = "OINF";
            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            myTemplate = doc.SelectSingleNode("/appSettings/configuration/Template").InnerText;
            CAName = doc.SelectSingleNode("/appSettings/configuration/CAName").InnerText;
            ExportPass = doc.SelectSingleNode("/appSettings/configuration/ExportPass").InnerText;
            string OUitem=doc.SelectSingleNode("/appSettings/configuration/OU").InnerText;
            this.comboBox1.Items.AddRange(OUitem.Split(','));          
            comboBox1.SelectedIndex = 0;
        }

  
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            X509Store store = new X509Store(StoreLocation.LocalMachine);
            store.Open(OpenFlags.MaxAllowed);

            X509Certificate2Collection CertColl = store.Certificates.Find(X509FindType.FindByIssuerName, CAName, true);
            if (CertColl.Count == 0)
            {
                //MessageBox.Show("No Certificate found!!!");
                label3.Text = "Certifikát nenalezen v uložišti.";
                Close();

            }
            else
            {
                foreach (X509Certificate2 Cert in CertColl)
                {
                    //MessageBox.Show(Cert.Subject);
                    if (Cert.Subject.Contains(textBox1.Text))
                    {
                        certData = Cert.Export(X509ContentType.Pkcs12, ExportPass);

                        
                        saveFileDialog1.Filter = "Pfx certifikát|*.pfx";
                        saveFileDialog1.Title = "Uložit certifikát...";
                        saveFileDialog1.FileName = textBox1.Text+".pfx";
                        saveFileDialog1.ShowDialog();
                        
                    }

                }
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //exportButton.Enabled = false;
            File.WriteAllBytes(saveFileDialog1.FileName, certData);
            label3.Text = "Export OK";
                       
        }
    }
}
