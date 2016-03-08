/*
* To change this license header, choose License Headers in Project Properties.
* To change this template file, choose Tools | Templates
* and open the template in the editor.
*/
package tenshelpingtool_1;
import com.google.api.client.auth.oauth2.TokenResponseException;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.auth.oauth2.GoogleTokenResponse;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson.JacksonFactory;
import com.google.gdata.client.spreadsheet.*;
import com.google.gdata.data.spreadsheet.*;
import com.google.gdata.util.*;
import static com.sun.org.apache.xalan.internal.xsltc.compiler.Constants.REDIRECT_URI;
import java.io.BufferedReader;

import java.io.IOException;
import java.io.InputStreamReader;
import java.net.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.Drive.Files;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.FileList;
import com.google.gdata.data.Link;
import com.google.gdata.data.batch.BatchOperationType;
import com.google.gdata.data.batch.BatchStatus;
import com.google.gdata.data.batch.BatchUtils;
import com.sun.deploy.util.SessionState.Client;
import gui.ava.html.image.generator.HtmlImageGenerator;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.MediaTracker;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.Toolkit;
import java.awt.Transparency;
import java.awt.image.BufferedImage;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import javax.swing.Icon;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.nio.charset.Charset;
import javax.imageio.ImageIO;
import javax.print.PrintService;
import javax.swing.JPanel;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import org.apache.commons.io.IOUtils;
import org.apache.fop.apps.FOPException;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Node;

import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.MimeConstants;
import org.apache.poi.hwpf.converter.WordToFoConverter;

import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.MimeConstants;
import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.http.HttpResponse;
/**
 *
 * @author Erfan
 */
public class TENsHelpingTool_1 extends javax.swing.JFrame {
    
    /**
     * Creates new form firstPage
     */
    private static String theTempDir = "b";
    private static String REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";
    private static String CLIENT_ID = "584057182641-2u253gnvkjnhnhqva49rm5e0ce2621a6.apps.googleusercontent.com";
    private static String CLIENT_SECRET = "jmrFxXdfzhKT2-FlEmixUWVJ";
    private  static List<SpreadsheetEntry> spreadsheets;
    private  static SpreadsheetEntry spreadsheet;
    static List<WorksheetEntry> worksheets;
    static WorksheetEntry worksheet;
    static URL listFeedUrl;
    private static SpreadsheetService spreadSheetService;
    private static GoogleCredential credential;
    private static GoogleAuthorizationCodeFlow flow;
    private static int rowNumber ;
    private static ListFeed listFeed;
    //variables for customer info
    private static  SpreadsheetEntry products_sheet;
    private static ArrayList<ArrayList<String>> products_and_categories = new ArrayList<ArrayList<String>>();
    private static String customer_name= "temp var";
    private static String customer_phone= "temp var";
    private static String customer_address= "temp var";
    private static String delivery_option= "temp var";
    private static String customer_product_catagory = "temp var";
    private static String customer_product_name= "temp var";
    private static String orderid = "temp";
    private static String customer_product_code= "temp var";
    private static String customer_product_size= "temp var";
    private static String customer_product_quantity= "temp var";
    private static String custoemr_email= "temp var";
    private static String customer_final_product_details= "temp var";
    private static String customer_comment= "temp var";
    private static String customer_product_price= "temp var";
    private static String customer_shipping_price= "temp var";
    private static String customer_pfinal_price= "temp var";
    private static int worksheetmaximum_row;
    private static String customer_product_description= "temp var";
    private static List <ListEntry> list_feed_customers ;
    private static List <ListEntry> list_feed_products ;
    private static String customer_emailAddress;
    private static String zip = "1207";
    private static String mainDir = "C:\\";
    private static JsonFactory jsonFactory = null;
    private static HttpTransport httpTransport = null;
    private static String ecourierApiUrl="http://ecourier.com.bd/api/?parcel=order&user_id=E4334&api_key=waer&api_secret=EeUvT&recipient_name=somename&recipient_mobile=01711XXXXXX&recipient_address=someaddress&recipient_zip=XXXX&delevery_timing=1&shipping_area=1&parcel_weight=1&product_price=0000\n" +
            "API";
    private static HWPFDocument doc = null;
    public TENsHelpingTool_1() {
        initComponents();
    }
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        row_number = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        display_info_button = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        jTextField_productPrice = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jTextField_shippingCharge = new javax.swing.JTextField();
        jLabel4 = new javax.swing.JLabel();
        jTextField_totalPrice = new javax.swing.JTextField();
        jLabel_name = new javax.swing.JLabel();
        jLabel_product = new javax.swing.JLabel();
        jLabel_productname = new javax.swing.JLabel();
        jLabel_Productcode = new javax.swing.JLabel();
        jLabel_Size = new javax.swing.JLabel();
        jLabel_Contactnumber = new javax.swing.JLabel();
        jLabel_email = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jTextField5 = new javax.swing.JTextField();
        jLabel14 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTextArea_address = new javax.swing.JTextArea();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTextArea_Productdescription = new javax.swing.JTextArea();
        jLabel15 = new javax.swing.JLabel();
        jButton_makeInnovoice = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        jTextArea_comment = new javax.swing.JTextArea();
        jLabel16 = new javax.swing.JLabel();
        jButton_addToECourier = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();
        jComboBox2 = new javax.swing.JComboBox();
        jLabel_DeliveryOption = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jTextField_quantity = new javax.swing.JTextField();
        jComboBox_products = new javax.swing.JComboBox();
        jButton2 = new javax.swing.JButton();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();
        jSeparator2 = new javax.swing.JSeparator();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        row_number.setName("row_number"); // NOI18N

        jLabel1.setText("Row no: ");

        display_info_button.setText("Display info");
        display_info_button.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                display_info_buttonActionPerformed(evt);
            }
        });

        jLabel2.setText("Product price:");

        jTextField_productPrice.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_productPriceActionPerformed(evt);
            }
        });

        jLabel3.setText("Shipping charge:");

        jLabel4.setText("Total Price:");

        jLabel_name.setText("Name");

        jLabel_product.setText("Product");

        jLabel_productname.setText("Product name");

        jLabel_Productcode.setText("Product code");

        jLabel_Size.setText("Size");

        jLabel_Contactnumber.setText("Contact number");

        jLabel_email.setText("email");

        jLabel13.setText("Zip");

        jLabel14.setText("Address");

        jTextArea_address.setColumns(20);
        jTextArea_address.setRows(5);
        jScrollPane1.setViewportView(jTextArea_address);

        jTextArea_Productdescription.setColumns(20);
        jTextArea_Productdescription.setRows(5);
        jScrollPane2.setViewportView(jTextArea_Productdescription);

        jLabel15.setText("Product description");

        jButton_makeInnovoice.setText("Make innovoice");
        jButton_makeInnovoice.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_makeInnovoiceActionPerformed(evt);
            }
        });

        jTextArea_comment.setColumns(20);
        jTextArea_comment.setRows(5);
        jScrollPane3.setViewportView(jTextArea_comment);

        jLabel16.setText("Comment");

        jButton_addToECourier.setText("Add to ecourier");
        jButton_addToECourier.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_addToECourierActionPerformed(evt);
            }
        });

        jButton4.setText("Print");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Item 1" }));
        jComboBox2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox2ActionPerformed(evt);
            }
        });

        jLabel_DeliveryOption.setText("Delivery Option");

        jLabel5.setText("Quantity:");

        jTextField_quantity.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_quantityActionPerformed(evt);
            }
        });

        jComboBox_products.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jButton2.setText("Prepare product");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jLabel6.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jLabel6.setText("For Processing Products");

        jLabel7.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jLabel7.setText("For Processing Invoice ");

        jSeparator2.setOrientation(javax.swing.SwingConstants.VERTICAL);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 193, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jComboBox_products, javax.swing.GroupLayout.Alignment.LEADING, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel6, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 198, Short.MAX_VALUE))
                    .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 193, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                        .addComponent(jButton_addToECourier, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(display_info_button, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 196, Short.MAX_VALUE)
                        .addGroup(layout.createSequentialGroup()
                            .addComponent(jLabel1)
                            .addGap(33, 33, 33)
                            .addComponent(row_number, javax.swing.GroupLayout.PREFERRED_SIZE, 96, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addComponent(jButton_makeInnovoice, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton4, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addComponent(jSeparator1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 208, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 11, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 23, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jLabel_product, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel5)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 92, Short.MAX_VALUE)
                        .addComponent(jTextField_quantity, javax.swing.GroupLayout.PREFERRED_SIZE, 184, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel_Size, javax.swing.GroupLayout.PREFERRED_SIZE, 304, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel_name, javax.swing.GroupLayout.PREFERRED_SIZE, 304, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel_productname, javax.swing.GroupLayout.PREFERRED_SIZE, 301, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel_Productcode, javax.swing.GroupLayout.PREFERRED_SIZE, 304, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel16)
                    .addComponent(jLabel15)
                    .addComponent(jScrollPane2))
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(43, 43, 43)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel_Contactnumber, javax.swing.GroupLayout.PREFERRED_SIZE, 304, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, 188, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                        .addComponent(jLabel_DeliveryOption, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel_email, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 291, Short.MAX_VALUE))
                                    .addComponent(jLabel14))
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING)))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 80, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jTextField_productPrice, javax.swing.GroupLayout.PREFERRED_SIZE, 144, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(jLabel4))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                    .addComponent(jTextField_totalPrice, javax.swing.GroupLayout.DEFAULT_SIZE, 144, Short.MAX_VALUE)
                                    .addComponent(jTextField_shippingCharge))))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(15, 15, 15)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel_name)
                            .addComponent(jLabel_Contactnumber))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel_product)
                            .addComponent(jLabel_email))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel_productname)
                            .addComponent(jLabel13)
                            .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel_DeliveryOption, javax.swing.GroupLayout.PREFERRED_SIZE, 49, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jLabel14)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jScrollPane1))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel_Productcode)
                                .addGap(10, 10, 10)
                                .addComponent(jLabel_Size, javax.swing.GroupLayout.PREFERRED_SIZE, 52, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel5)
                                    .addComponent(jTextField_quantity, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jLabel15)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 113, javax.swing.GroupLayout.PREFERRED_SIZE))))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(23, 23, 23)
                        .addComponent(jLabel6)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jComboBox_products, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jButton2)
                        .addGap(37, 37, 37)
                        .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jLabel7)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(56, 56, 56)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 14, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(row_number, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addComponent(display_info_button))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(15, 15, 15)
                        .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 133, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel16)
                            .addComponent(jButton_makeInnovoice))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 13, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 53, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jButton4)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jButton_addToECourier))))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField_productPrice, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField_shippingCharge, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jTextField_totalPrice, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addGap(10, 10, 10))
        );

        row_number.getAccessibleContext().setAccessibleName("row_number");

        pack();
    }// </editor-fold>//GEN-END:initComponents
    
    private void display_info_buttonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_display_info_buttonActionPerformed
        JFrame frame = new JFrame("invalid number");
        int readyToGoWithRowNumberFlag =0;
        try{
            rowNumber= Integer.parseInt(row_number.getText());
            if(rowNumber>2 && rowNumber < worksheetmaximum_row){
                readyToGoWithRowNumberFlag++;
            }else{
                readyToGoWithRowNumberFlag =0;
                JOptionPane.showMessageDialog(frame, "Not a valid row number!");
            }
        }catch(Exception e){
            readyToGoWithRowNumberFlag =0;
            JOptionPane.showMessageDialog(frame, "Not a valid row number!");
        }
        if(readyToGoWithRowNumberFlag > 0){
            fillInCustomerInfo();
        }
        jTextField_productPrice.setText("0");
    }//GEN-LAST:event_display_info_buttonActionPerformed
    
    private void jComboBox2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox2ActionPerformed
        if(jComboBox2.getSelectedItem() != null && !jComboBox2.getSelectedItem().equals("item1")){
            spreadSheetSelect();
        }
    }//GEN-LAST:event_jComboBox2ActionPerformed
    
    private void jTextField_quantityActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_quantityActionPerformed
        // TODO add your handling code here:
        customer_product_quantity = jTextField_quantity.getText();
        make_description();
    }//GEN-LAST:event_jTextField_quantityActionPerformed
    
    private void jButton_makeInnovoiceActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_makeInnovoiceActionPerformed
        // TODO add your handling code here:
        makeOrUpdateInvoice();
    }//GEN-LAST:event_jButton_makeInnovoiceActionPerformed
    
    private void jTextField_productPriceActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_productPriceActionPerformed
        // TODO add your handling code here:
        try{
            jTextField_totalPrice.setText(String.valueOf(Integer.parseInt(jTextField_shippingCharge.getText()) + Integer.parseInt(jTextField_productPrice.getText())));
        } catch (java.lang.NumberFormatException nfe){
            JFrame frame = new JFrame("invalid price");
            JOptionPane.showMessageDialog(frame, "The price you entered is not valid, please try again!");
        }
        customer_product_price = jTextField_productPrice.getText();
        
    }//GEN-LAST:event_jTextField_productPriceActionPerformed
    private org.apache.http.client.CookieStore m_cookies = null;    //cookie container ready to receive the forms auth cookie
    
    private void jButton_addToECourierActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_addToECourierActionPerformed
        System.out.println("adding information to ecourier");
        try{
            try{
                ecourierApiUrl = "http://ecourier.com.bd/api/?parcel=order&user_id=E4334&api_key=waer&api_secret=EeUvT&recipient_name="
                        +URLEncoder.encode(customer_name, "UTF-8")+
                        "&recipient_mobile=0"
                        +URLEncoder.encode(customer_phone, "UTF-8")+
                        "&recipient_address="
                        +URLEncoder.encode(customer_address, "UTF-8")+
                        "&recipient_zip="
                        +URLEncoder.encode(zip, "UTF-8")+
                        "&delevery_timing=2&shipping_area=1&parcel_weight=1&product_price="
                        +URLEncoder.encode(customer_product_price, "UTF-8");
                System.out.println(ecourierApiUrl);
                
            } catch (IOException ex) {
                Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
            }
            
            HttpClient client = new DefaultHttpClient();
            
            HttpGet request = new HttpGet(ecourierApiUrl);
            org.apache.http.HttpResponse response = client.execute(request);
            BufferedReader rd = new BufferedReader (new InputStreamReader(response.getEntity().getContent()));
            String line = "";
            while ((line = rd.readLine()) != null) {
                System.out.println(line);
            }
            /*
            
            DefaultHttpClient httpClient = new DefaultHttpClient();
            HttpPost httppost = new HttpPost(ecourierApiUrl);
            httppost.setHeader("Content-Type", "application/x-www-form-urlencoded");
            httppost.setHeader("Charset", "UTF-8");
            httppost.setHeader("Accept", "application/json");
            
            //access the secured service with the authorization cookie
            httpClient.setCookieStore(m_cookies);
            
            
            
            org.apache.http.HttpResponse response = null;
            String res = "";
            try {
            response= httpClient.execute(httppost);
            //res = (response.getEntity().getContent());
            String encoding ="UTF-8";
            
            res = IOUtils.toString(response.getEntity().getContent(), encoding);
            
            System.out.println("Response:" + res);
            } catch (org.apache.http.conn.HttpHostConnectException e) {
            System.out.println("Cannot receive response due to " + e);
            //if (url.contains(LOGIN_URL))
            //	res = TIMEOUT_RESP;
            } catch (IOException e) {
            System.out.println("Cannot parse incoming response due to " + e);
            } finally {
            httpClient.getConnectionManager().shutdown();
            }
            
            //save cookies for the next secured run:
            m_cookies = httpClient.getCookieStore();
            */
        } catch (IOException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jButton_addToECourierActionPerformed
    
    
    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        // TODO add your handling code here:
        ArrayList<String> SelectedProduct = new ArrayList<String>();
        String product_string = jComboBox_products.getSelectedItem().toString();
        for(int idx =0; idx < products_and_categories.size(); idx ++){
            ArrayList<String> temp = products_and_categories.get(idx);
            if(temp.get(0).equals(product_string)){
                SelectedProduct = temp;
            }
        }
        prepare_products(SelectedProduct);
    }//GEN-LAST:event_jButton2ActionPerformed
    
    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        try {
            WordToFoConverter wf = new WordToFoConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            wf.processDocument(doc);
            
            org.w3c.dom.Document foinvoice = wf.getDocument();
            
            FileWriter outaa = new FileWriter(theTempDir+"\\gu.fo");
            DOMSource domSource = new DOMSource( (Node) foinvoice );
            StreamResult streamResult = new StreamResult( outaa );
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            // TODO set encoding from a command argument
            serializer.setOutputProperty( OutputKeys.ENCODING, "UTF-8" );
            serializer.setOutputProperty( OutputKeys.INDENT, "yes" );
            serializer.transform( domSource, streamResult );
            outaa.close();
            FopFactory fopFactory = FopFactory.newInstance();
            OutputStream out = new BufferedOutputStream(new FileOutputStream(new File(theTempDir+"\\myfile.pdf")));
            
            // Step 3: Construct fop with desired output format
            Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, out);
            
            // Step 4: Setup JAXP using identity transformer
            TransformerFactory factory = TransformerFactory.newInstance();
            Transformer transformer = factory.newTransformer(); // identity transformer
            
            // Step 5: Setup input and output for XSLT transformation
            // Setup input stream
            //Source src = new StreamSource(new File("C:\\myfile.fo"));
            Source src = new StreamSource(new File(mainDir+"TENS Posters\\Temp\\gu.fo"));
            
            // Resulting SAX events (the generated FO) must be piped through to FOP
            Result res = new SAXResult(fop.getDefaultHandler());
            
            // Step 6: Start XSLT transformation and FOP processing
            transformer.transform(src, res);
            out.close();
            
            printPDF(theTempDir+"\\myfile.pdf", choosePrinter());
            
        }           catch (FileNotFoundException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (FOPException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerConfigurationException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }
        /*try {
        File tempDir = new File(mainDir+"\\TENS Posters\\Temp");
        // if the directory does not exist, create it
        if (!tempDir.exists()) {
        //System.out.println("creating directory: " + directoryName);
        boolean result = false;
        try{
        tempDir.mkdir();
        result = true;
        } catch(SecurityException se){
        //handle it
        }
        if(result) {
        System.out.println("Temp dir created");
        }
        }
        // TODO add your handling code here:
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
        DocumentBuilderFactory.newInstance().newDocumentBuilder()
        .newDocument());
        wordToHtmlConverter.processDocument(doc);
        org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource((Node) htmlDocument);
        StreamResult streamResult = new StreamResult(out);
        
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-16");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        System.out.println(out);
        out.close();
        String tempHtmlSvaeDir = mainDir+"\\TENS Posters\\Temp\\tensinvoice.html";
        OutputStream outputStream = new FileOutputStream (tempHtmlSvaeDir);
        out.writeTo(outputStream);
        System.out.println("Html saved in a temp folder successfully");
        HtmlImageGenerator imageGenerator = new HtmlImageGenerator();
        //String tempHtml =  new String( out.toByteArray(), "UTF-8" );
        //System.out.println(tempHtml);
        imageGenerator.loadUrl("file:"+tempHtmlSvaeDir);
        System.out.println("html loaded successfulyl!");
        //Dimension d = new Dimension(1654 ,2339);
        imageGenerator.getBufferedImage();
        
        
        } catch (ParserConfigurationException ex) {
        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerConfigurationException ex) {
        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerException ex) {
        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }*/
        catch (ParserConfigurationException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (PrinterException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }//GEN-LAST:event_jButton4ActionPerformed
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) throws IOException, MalformedURLException, ServiceException{
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
        * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html
        */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(TENsHelpingTool_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(TENsHelpingTool_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(TENsHelpingTool_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(TENsHelpingTool_1.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        
        Oauth2();
        tempDirCreation();
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new TENsHelpingTool_1().setVisible(true);
                connect();
            }
        });
        
    }
    private static void  connect(){
        JFrame frame = new JFrame("Connection stuff");
        try {
            spreadSheetService =new SpreadsheetService("MySpreadsheetIntegration-v1");
            spreadSheetService.setOAuth2Credentials(credential);
            
            // Define the URL to request.  This should never change.
            URL SPREADSHEET_FEED_URL = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = spreadSheetService.getFeed(SPREADSHEET_FEED_URL, SpreadsheetFeed.class);
            spreadsheets = feed.getEntries();
            jComboBox2.removeAllItems();
            jComboBox_products.removeAllItems();
            for (SpreadsheetEntry spreadsheet : spreadsheets) {
                String temp = spreadsheet.getTitle().getPlainText();
                if(temp.contains("productcategories")){
                    products_sheet = spreadsheet;
                    System.out.println(products_sheet.getTitle().getPlainText());
                    worksheets = products_sheet.getWorksheets();
                    worksheet = worksheets.get(0);
                    listFeedUrl = worksheet.getListFeedUrl();
                    listFeed = spreadSheetService.getFeed(listFeedUrl, ListFeed.class);
                    
                    list_feed_customers = listFeed.getEntries();
                    ListEntry row = list_feed_customers.get(2);
                    for (String tag : row.getCustomElements().getTags()) {
                        jComboBox_products.addItem(tag);
                        ArrayList<String> tempa = new ArrayList<String>();
                        tempa.add(tag);
                        products_and_categories.add(tempa);
                    }
                    for(int aw=0; aw<list_feed_customers.size();aw++){
                        ListEntry rowa = list_feed_customers.get(aw);
                        for(int g=0; g< products_and_categories.size(); g++){
                            //added later 1st dec 2014
                            String product =products_and_categories.get(g).get(0);
                            String subCat = rowa.getCustomElements().getValue(product);
                            if(subCat !=null){
                                for(int awo=0; awo<list_feed_customers.size();awo++){
                                    if(products_and_categories.get(awo).get(0).equals(product)){
                                        products_and_categories.get(awo).add(subCat);
                                    }
                                }
                            }
                            //till here
                        }
                    }
                }
            }
            for (SpreadsheetEntry spreadsheet : spreadsheets) {
                String temp = spreadsheet.getTitle().getPlainText();
                if(temp.contains("(ORSS)") || temp.equals("Orders of Nemesis (Responses)")|| temp.equals("Place an order ( HOMICIDE ) (Responses)")){
                    jComboBox2.addItem(temp);
                }
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(frame, "There was an io exception!. \n The programm will exit now...");
            System.exit(0);
        } catch (ServiceException ex) {
            JOptionPane.showMessageDialog(frame, "There was a service exception \n The programm will exit now...");
            System.exit(0);
        }
    }
    private static void spreadSheetSelect(){
        try {
            String temp2;
            for(int i =0; i < spreadsheets.size();i++){
                temp2 = spreadsheets.get(i).getTitle().getPlainText();
                if(temp2.equals(jComboBox2.getSelectedItem().toString())){
                    spreadsheet = spreadsheets.get(i);
                    System.out.println(spreadsheet.getTitle().getPlainText());
                    break;
                }
            }
            System.out.println(spreadsheet.getTitle().getPlainText());
            worksheets = spreadsheet.getWorksheets();
            worksheet = worksheets.get(0);
            listFeedUrl = worksheet.getListFeedUrl();
            listFeed = spreadSheetService.getFeed(listFeedUrl, ListFeed.class);
            
            list_feed_customers = listFeed.getEntries();
            worksheetmaximum_row = list_feed_customers.size()+2;
            System.out.println("maximum row: "+ worksheetmaximum_row);
        } catch (IOException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ServiceException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    private static void prepare_products(ArrayList<String> p){
            File tensDir = new File(mainDir+"\\TENS Posters");
            // if the directory does not exist, create it
            if (!tensDir.exists()) {
                //System.out.println("creating directory: " + directoryName);
                boolean result = false;
                try{
                    tensDir.mkdir();
                    result = true;
                } catch(SecurityException se){
                    //handle it
                }
                if(result) {
                    System.out.println("pre-preparing tasks: TENS Posters DIR created");
                }
            }
            File productsDir = new File(mainDir+"\\TENS Posters\\Prepared Products");
            // if the directory does not exist, create it
            if (!productsDir.exists()) {
                //System.out.println("creating directory: " + directoryName);
                boolean result = false;
                try{
                    productsDir.mkdir();
                    result = true;
                } catch(SecurityException se){
                    //handle it
                }
                if(result) {
                    System.out.println("pre-preparing tasks: Products dir created");
                }
            }
            System.out.println("pre-preparing info: product tobe prepared: "+ p.get(0));
            String product_type = p.get(0);
            String prep_method = p.get(1);
            File productsTypeDir = new File(mainDir+"\\TENS Posters\\Prepared Products\\"+product_type);
            // if the directory does not exist, create it
            if (!productsTypeDir.exists()) {
                //System.out.println("creating directory: " + directoryName);
                boolean result = false;
                try{
                    productsTypeDir.mkdir();
                    result = true;
                } catch(SecurityException se){
                    //handle it
                }
                if(result) {
                    System.out.println("pre-preparing tasks: Products type dir created");
                }
            }
            //Create a new authorized API client
            Drive driveService = new Drive.Builder(httpTransport, jsonFactory, credential).build();
            System.out.println("pre-preparing tasks: Drive service building done!");
            
            //debug purpose
            int counter =0;
            
            ArrayList<ArrayList<String>> processed_product_code_quantity= new ArrayList<ArrayList<String>>();
            for(int rowIdx =1; rowIdx < list_feed_customers.size(); rowIdx++){
                ListEntry row_prep = list_feed_customers.get(rowIdx);
                String productCode = null;
                String productQuantity = null;
                String productSize = null;
                if(row_prep.getCustomElements().getValue("status").equals("0")){
                    //System.out.println("going to find product code");
                    for (String tag : row_prep.getCustomElements().getTags()) {
                        String tag_lower = tag.toString().toLowerCase().replaceAll("\\s+","");
                        //System.out.println("tag: "+ tag_lower);
                        if(tag_lower.contains(product_type.toLowerCase())){
                            //System.out.println("found the product "+ product_type);
                            if(tag_lower.contains("code")){
                                productCode = row_prep.getCustomElements().getValue(tag);//(row_prep.getCustomElements().getValue(tag) !=null) ? row_prep.getCustomElements().getValue(tag).replaceAll("^\"|\"$", "") : null;
                                
                            }
                            if(tag_lower.contains("quantity")){
                                productQuantity = row_prep.getCustomElements().getValue(tag);//(row_prep.getCustomElements().getValue(tag) !=null) ? row_prep.getCustomElements().getValue(tag).replaceAll("^\"|\"$", "") : null;
                                
                            }
                            if(tag_lower.contains("size")){
                                productSize = row_prep.getCustomElements().getValue(tag);//(row_prep.getCustomElements().getValue(tag) !=null) ? row_prep.getCustomElements().getValue(tag).replaceAll("^\"|\"$", "") : null;
                                
                            }                            
                        }
                    }
                }
                
                if(productCode != null){
                    productCode = productCode.replaceAll("\\s", "");
                    System.out.println("product Code: " +productCode+ " quantity: "+productQuantity+ " size: "+productSize);
                    //
                    
                    if(!productCode.contains("custom")){
                        if(productCode.contains(",")){
                            String[] codes = productCode.split(",");
                            String[] quantities = productQuantity.split(",");
                            if(codes.length==quantities.length){
                                
                                for(int bla=0; bla<codes.length; bla++){
                                    if(processed_product_code_quantity.size() ==0){
                                        //System.out.println("checkpoint!");
                                        ArrayList<String> temppoo = new ArrayList<String>();
                                        temppoo.add(codes[bla]);
                                        temppoo.add(quantities[bla]);
                                        temppoo.add(productSize);
                                        processed_product_code_quantity.add(temppoo);
                                    }else{
                                        int flag =0;
                                        for(int lmidx=0; lmidx < processed_product_code_quantity.size(); lmidx++){
                                            if(processed_product_code_quantity.size() !=0 && processed_product_code_quantity.get(lmidx).get(0).equals(codes[bla]) && processed_product_code_quantity.get(lmidx).get(2).equals(productSize)){
                                                //System.out.println("checkpoint!");
                                                processed_product_code_quantity.get(lmidx).set(1, String.valueOf(Integer.parseInt(processed_product_code_quantity.get(lmidx).get(1)) + Integer.parseInt(quantities[bla])));
                                                flag++;
                                                break;
                                            }
                                        }
                                        if(flag ==0){
                                            ArrayList<String> temppoo = new ArrayList<String>();
                                            temppoo.add(codes[bla]);
                                            temppoo.add(quantities[bla]);
                                            temppoo.add(productSize);
                                            processed_product_code_quantity.add(temppoo);
                                        }
                                    }
                                }
                            }
                        }else{
                            if(processed_product_code_quantity.size() ==0){
                                //System.out.println("checkpoint! starting from zero");
                                ArrayList<String> temppoo = new ArrayList<String>();
                                temppoo.add(productCode);
                                temppoo.add(productQuantity);
                                temppoo.add(productSize);
                                processed_product_code_quantity.add(temppoo);
                            }else{
                                int flag =0;
                                for(int lmidx=0; lmidx < processed_product_code_quantity.size(); lmidx++){
                                    if(processed_product_code_quantity.get(lmidx).get(0).equals(productCode) && processed_product_code_quantity.get(lmidx).get(2).equals(productSize)){
                                        System.out.println("checkpoint! repeat...null pointer pera");
                                        processed_product_code_quantity.get(lmidx).set(1, String.valueOf(Integer.parseInt(processed_product_code_quantity.get(lmidx).get(1)) + Integer.parseInt(productQuantity)));
                                        flag++;
                                        break;
                                    }
                                }
                                if (flag ==0){
                                    ArrayList<String> temppoo = new ArrayList<String>();
                                    temppoo.add(productCode);
                                    temppoo.add(productQuantity);
                                    temppoo.add(productSize);
                                    processed_product_code_quantity.add(temppoo);
                                }
                            }
                        }
                    }else{
                        //custom handeling
                        
                    }
                }
            }
            
            System.out.println("before downloading designs: files been listed in the prog, now lets test the progs storage...size of list:" +processed_product_code_quantity.size());
            for (int gu =0; gu< processed_product_code_quantity.size();gu++){
                for(int gi=0; gi< processed_product_code_quantity.get(gu).size(); gi++){
                    System.out.print(processed_product_code_quantity.get(gu).get(gi) + "  ");
                }
                System.out.println();
            }
            System.out.println("before downloading designs: done testing the progs storage...");
            if(prep_method.equals("0")){
                //rename method
                
                System.out.println("The product " + p.get(0) + "is selected to be processed in type "+ p.get(1) + " and list size is " + processed_product_code_quantity.size());
                for(int iii = 0; iii<processed_product_code_quantity.size(); iii++){
                    //mainDir+"\\TENS Posters\\Prepared Products\\"+product_type
                    String renam_method_dir = mainDir+"\\TENS Posters\\Prepared Products\\"+product_type;
                    File productsTypeCatDir = new File(renam_method_dir);
                    // if the directory does not exist, create it
                    if (!productsTypeCatDir.exists()) {
                        //System.out.println("creating directory: " + directoryName);
                        boolean result = false;
                        try{
                            productsTypeCatDir.mkdir();
                            result = true;
                        } catch(SecurityException se){
                            //handle it
                        }
                        if(result) {
                            System.out.println("Products type catagory dir created");
                        }
                    }
                    
                    try {
                        Files.List request = driveService.files().list()
                                .setQ("(mimeType = 'image/jpeg' or mimeType = 'image/png') and title contains  '" + processed_product_code_quantity.get(iii).get(0) + "' and trashed = false");
                        //debug printline
                        System.out.println("file.list ready for: "+processed_product_code_quantity.get(iii).get(0));
                        FileList files = request.execute();
                        System.out.println("request executed!");
                        if (files != null) {
                            for (com.google.api.services.drive.model.File file : files.getItems()) {
                                System.out.println("downloading file from gdrive..."+ processed_product_code_quantity.get(iii).get(0));
                                InputStream in = downloadFile(driveService, file);
                                String fileType = file.getMimeType();
                                if(fileType.contains("image/")){
                                        String tempyio = fileType.substring(6, fileType.length());
                                        fileType = "."+tempyio;
                                }
                                System.out.println("file is now in inputstream.. and type is: "+fileType);
                                String sizeWithoutPriceTag = processed_product_code_quantity.get(iii).get(2);
                                int idxOfBrack=0;
                                for(int oii=0; oii<sizeWithoutPriceTag.length();oii++){
                                    if(sizeWithoutPriceTag.charAt(oii)=='('){
                                        idxOfBrack = oii;
                                    }
                                }
                                String tempk = sizeWithoutPriceTag.substring(0, idxOfBrack);
                                sizeWithoutPriceTag = tempk;
                                String quantityy = processed_product_code_quantity.get(iii).get(1);
                                String f = renam_method_dir+"\\" + sizeWithoutPriceTag+"_quantity "+quantityy+fileType;
                                System.out.println("file dir and name: " + f);
                                byte[] buffer = new byte[8 * 1024];

try {
    
  OutputStream output = new FileOutputStream(f);
  System.out.println("outputstream created..");
  try {
    int bytesRead;
    while ((bytesRead = in.read(buffer)) != -1) {
      output.write(buffer, 0, bytesRead);
    }
    System.out.println("done writing output..");
  } finally {
    output.close();
  }
} finally {
  in.close();
}
                                if(processed_product_code_quantity.get(iii).size() > 3){
                                }else{
                                    processed_product_code_quantity.get(iii).add("success");
                                }
                            }
                            
                        }
                        if(processed_product_code_quantity.get(iii).size() == 4){                        
                            System.out.println(processed_product_code_quantity.get(iii).get(3));
                        }
                    } catch (IOException ex) {
                        //Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (Exception ex) {
//                        /Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
                    } 
                }
            }else if(prep_method.contains("1")){
                //folder wise image edit
                System.out.println("The product " + p.get(0) + "is selected to be processed in type "+ p.get(1) + " and list size is " + processed_product_code_quantity.size());
                for(int iii = 0; iii<processed_product_code_quantity.size(); iii++){
                    //mainDir+"\\TENS Posters\\Prepared Products\\"+product_type
                    File productsTypeCatDir = new File(mainDir+"\\TENS Posters\\Prepared Products\\"+product_type+"\\"+processed_product_code_quantity.get(iii).get(2));
                    // if the directory does not exist, create it
                    if (!productsTypeCatDir.exists()) {
                        //System.out.println("creating directory: " + directoryName);
                        boolean result = false;
                        try{
                            productsTypeCatDir.mkdir();
                            result = true;
                        } catch(SecurityException se){
                            //handle it
                        }
                        if(result) {
                            System.out.println("Products type catagory dir created");
                        }
                    }
                    
                    try {
                        Files.List request = driveService.files().list()
                                .setQ("(mimeType = 'image/jpeg' or mimeType = 'image/png') and title contains  '" + processed_product_code_quantity.get(iii).get(0) + "' and trashed = false");
                        //debug printline
                        System.out.println("file.list ready for: "+processed_product_code_quantity.get(iii).get(0));
                        FileList files = request.execute();
                        System.out.println("request executed!");
                        if (files != null) {
                            for (com.google.api.services.drive.model.File file : files.getItems()) {
                                System.out.println("downloading file..."+ processed_product_code_quantity.get(iii).get(0));
                                InputStream in = downloadFile(driveService, file);
                                System.out.println("resizing.." + processed_product_code_quantity.get(iii).get(0));
                                File theFile = resizeIt(stream2file(in),400,400);
                                System.out.println("writing pcs at the corner of the file.." + processed_product_code_quantity.get(iii).get(0));
                                BufferedImage in3 = ImageIO.read(theFile);
                                Graphics2D graph = in3.createGraphics();
                                graph.setColor(Color.BLACK);
                                graph.fill(new Rectangle(340,10 , 100, 30));
                                graph.setFont(graph.getFont().deriveFont(20f));
                                graph.setColor(Color.WHITE);
                                graph.drawString(processed_product_code_quantity.get(iii).get(1)+" pcs", 340, 30);
                                graph.dispose();
                                
                                System.out.println("saving file " + processed_product_code_quantity.get(iii).get(0));
                                String f = productsTypeCatDir+"\\" + processed_product_code_quantity.get(iii).get(1)+" "+file.getTitle();
                                ImageIO.write(in3, "png", new File(f));
                                in.close();
                                if(processed_product_code_quantity.get(iii).size() > 3){
                                }else{
                                    processed_product_code_quantity.get(iii).add("success");
                                }
                            }
                            
                        }
                        if(processed_product_code_quantity.get(iii).size() == 4){                        
                            System.out.println(processed_product_code_quantity.get(iii).get(3));
                        }
                    } catch (IOException ex) {
                        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (Exception ex) {
                        Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            }
            
            for(int pl=0; pl<processed_product_code_quantity.size();pl++){
                for(int pll=0; pll<processed_product_code_quantity.get(pl).size();pll++){
                    if(processed_product_code_quantity.get(pl).size()<4){
                        processed_product_code_quantity.get(pl).add("failed");
                    }
                }
            }
            for(int pl=0; pl<processed_product_code_quantity.size();pl++){
                for(int pll=0; pll<processed_product_code_quantity.get(pl).size();pll++){
                    System.out.print(processed_product_code_quantity.get(pl).get(pll)+" ");
                }
                System.out.println();
            }
            //now lets update the cells
            /*the reason of updating the cell values later is because the downloads may take
            time and in the meant ime some one can place and order of the same product and to properly count the
            products cells are being updated later on depending on the downloaded files
            */
            List<String> cellsToBeUpdated= new ArrayList<String>();
            for(int rowIdx =1; rowIdx < list_feed_customers.size(); rowIdx++){
                String productCode = null;
                String productQuantity = null;
                ListEntry row_prep = list_feed_customers.get(rowIdx);
                if(row_prep.getCustomElements().getValue("status").equals("0")){
                    //System.out.println("going to find product code");
                    for (String tag : row_prep.getCustomElements().getTags()) {
                        String tag_lower = tag.toString().toLowerCase().replaceAll("\\s+","");
                        //System.out.println("tag: "+ tag_lower);
                        if(tag_lower.contains(product_type.toLowerCase())){
                            //System.out.println("found the product "+ product_type);
                            if(tag_lower.contains("code")){
                                productCode = row_prep.getCustomElements().getValue(tag);//(row_prep.getCustomElements().getValue(tag) !=null) ? row_prep.getCustomElements().getValue(tag).replaceAll("^\"|\"$", "") : null;
                                
                            }
                            if(tag_lower.contains("quantity")){
                                productQuantity = row_prep.getCustomElements().getValue(tag);//(row_prep.getCustomElements().getValue(tag) !=null) ? row_prep.getCustomElements().getValue(tag).replaceAll("^\"|\"$", "") : null;
                                
                            }                            
                        }
                    }
                }
                                            
                for(int guio=0; guio<processed_product_code_quantity.size(); guio++){
                    if(processed_product_code_quantity.get(guio).get(0).equals(productCode) && processed_product_code_quantity.get(guio).get(3).equals("success")){
                        if(Integer.parseInt(productQuantity) <= Integer.parseInt(processed_product_code_quantity.get(guio).get(1))){
                            processed_product_code_quantity.get(guio).set(2, String.valueOf(Integer.parseInt(processed_product_code_quantity.get(guio).get(1)) - Integer.parseInt(productQuantity)));
                            cellsToBeUpdated.add(String.valueOf(rowIdx+2));
                        }
                    }
                }
            }
            URL cellFeedUrl = worksheet.getCellFeedUrl();
            System.out.println("cellfeedurl has been prepared!");
        try {
            CellFeed cellFeed = spreadSheetService.getFeed(cellFeedUrl, CellFeed.class);
            System.out.println("cell feed has been prepared!");
            for(String o: cellsToBeUpdated){
                System.out.println("feedback list's row number: "+ o);
                int counterr = 0;
            for (CellEntry cell : cellFeed.getEntries()) {
                String theCell = "B"+o;counterr++;
                System.out.println("cell checking: "+theCell+ "  current cell : "+cell.getTitle().getPlainText());
                if (cell.getTitle().getPlainText().equals(theCell)) {
                    System.out.println("Should be updating "+theCell+" cell to 1ss");
                cell.changeInputValueLocal("1ss");
                cell.update();
                break;
            }
            }
            System.out.println("cell feed updating: done "+counterr);
        }
            
        } catch (IOException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ServiceException ex) {
            Logger.getLogger(TENsHelpingTool_1.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    public static File resizeIt(File file, int width, int height) throws Exception{
        Image img = Toolkit.getDefaultToolkit().getImage( file.getAbsolutePath() );
        loadCompletely(img);
        BufferedImage bm = toBufferedImage(img);
        bm = resize(bm, width, height);
        
        StringBuilder sb = new StringBuilder();
        sb.append( bm.hashCode() ).append(".png");
        String filename = sb.toString();
        File result = new File( filename );
        ImageIO.write(bm, "png", result);
        return result;
    }
    private static BufferedImage toBufferedImage(Image img){
        if (img instanceof BufferedImage){
            return (BufferedImage) img;
        }
        
        BufferedImage bimage = new BufferedImage(img.getWidth(null), img.getHeight(null), BufferedImage.TYPE_INT_ARGB);
        
        bimage.getGraphics().drawImage(img, 0, 0 , null);
        bimage.getGraphics().dispose();
        
        return bimage;
    }
    
    private static BufferedImage resize(BufferedImage image, int areaWidth, int areaHeight){
        float scaleX = (float) areaWidth / image.getWidth();
        float scaleY = (float) areaHeight / image.getHeight();
        float scale = Math.min(scaleX, scaleY);
        int w = Math.round(image.getWidth() * scale);
        int h = Math.round(image.getHeight() * scale);
        int type = image.getTransparency() == Transparency.OPAQUE ? BufferedImage.TYPE_INT_RGB : BufferedImage.TYPE_INT_ARGB;
        boolean scaleDown = scale < 1;
        if (scaleDown) {
            // multi-pass bilinear div 2
            int currentW = image.getWidth();
            int currentH = image.getHeight();
            BufferedImage resized = image;
            while (currentW > w || currentH > h) {
                currentW = Math.max(w, currentW / 2);
                currentH = Math.max(h, currentH / 2);
                BufferedImage temp = new BufferedImage(currentW, currentH, type);
                Graphics2D g2 = temp.createGraphics();
                g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                g2.drawImage(resized, 0, 0, currentW, currentH, null);
                g2.dispose();
                resized = temp;
            }
            return resized;
        } else {
            Object hint = scale > 2 ? RenderingHints.VALUE_INTERPOLATION_BICUBIC : RenderingHints.VALUE_INTERPOLATION_BILINEAR;
            BufferedImage resized = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2 = resized.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, hint);
            g2.drawImage(image, 0, 0, w, h, null);
            g2.dispose();
            return resized;
        }
    }
    private static void loadCompletely (Image img){
        MediaTracker tracker = new MediaTracker(new JPanel());
        tracker.addImage(img, 0);
        try {
            tracker.waitForID(0);
        } catch (InterruptedException ex) {
            throw new RuntimeException(ex);
        }
    }
    private static final String PREFIX = "stream2file";
    private static final String SUFFIX = ".tmp";
    
    private static InputStream downloadFile(Drive service, com.google.api.services.drive.model.File file) {
        if (file.getDownloadUrl() != null && file.getDownloadUrl().length() > 0) {
            try {
                HttpResponse resp =
                        service.getRequestFactory().buildGetRequest(new GenericUrl(file.getDownloadUrl()))
                                .execute();
                return resp.getContent();
            } catch (IOException e) {
                // An error occurred.
                e.printStackTrace();
                return null;
            }
        } else {
            // The file doesn't have any content stored on Drive.
            return null;
        }
    }
    public static File stream2file (InputStream in) throws IOException {
        final File tempFile = File.createTempFile(PREFIX, SUFFIX);
        tempFile.deleteOnExit();
        try (FileOutputStream out = new FileOutputStream(tempFile)) {
            IOUtils.copy(in, out);
        }
        return tempFile;
    }
    private static void Oauth2() throws IOException{
        httpTransport = new NetHttpTransport();
        jsonFactory = new JacksonFactory();
        List <String> SCOPES= new ArrayList(Arrays.asList(DriveScopes.DRIVE));
        SCOPES.addAll(Arrays.asList("https://spreadsheets.google.com/feeds"));
        String gmailSCOPE = "https://www.googleapis.com/auth/gmail.readonly";
        SCOPES.add(gmailSCOPE);
        flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, jsonFactory, CLIENT_ID, CLIENT_SECRET, SCOPES)
                .setAccessType("online")
                .setApprovalPrompt("auto")
                .build();
        
        String url = flow.newAuthorizationUrl().setRedirectUri(REDIRECT_URI).build();
        //test 2
        JFrame frame = new JFrame("Oauth2 authorization");
        String code = JOptionPane.showInputDialog(
                frame,
                "<html><body><p style='width: 300px;'> <h4>Please open the following URL in your browser to get the authorization code:</h4> <textarea rows=\"1\" cols=\"50\" readonly > "+ url +"</textarea> </br> <h4>enter the authorization code here:</h4></body></html>",
                "authorization code needed",
                JOptionPane.WARNING_MESSAGE
        );
        System.out.printf("The secret code is '%s'.\n", code);
        try{
            GoogleTokenResponse response = flow.newTokenRequest(code).setRedirectUri(REDIRECT_URI).execute();
            credential= new GoogleCredential().setFromTokenResponse(response);
            System.out.println("oauth2.0 credentials done!");
        }catch (TokenResponseException e){
            JOptionPane.showMessageDialog(frame, "The authorization code was not correct. \n The programm will exit now...");
            System.exit(0);
        }
        catch(java.lang.NullPointerException ne){
            JOptionPane.showMessageDialog(frame, "The authorization process was cancelled. \n The programm will exit now...");
            System.exit(0);
        }catch (java.net.UnknownHostException uhe){
            JOptionPane.showMessageDialog(frame, "Something must be wrong with your internet . \n The programm will exit now...");
            System.exit(0);
        }
    }
    private static void  fillInCustomerInfo(){
        ListEntry row = list_feed_customers.get(rowNumber-2);
        for (String tag : row.getCustomElements().getTags()) {
            String temp = tag.toString().toLowerCase().replaceAll("\\s+","");
            System.out.println("tag: "+ temp);
            if(temp.equals("name")){
                customer_name = row.getCustomElements().getValue(tag);
                jLabel_name.setText("Name: " + customer_name);
                System.out.println("customer name: " + customer_name);
            } if(temp.equals("emailaddress")){
                customer_emailAddress = row.getCustomElements().getValue(tag);
                jLabel_email.setText("email: "+ customer_emailAddress);
                System.out.println("customer_emailAddress: "+ customer_emailAddress);
            }if(temp.equals("contactnumber")){
                customer_phone = row.getCustomElements().getValue(tag);
                jLabel_Contactnumber.setText("Contact no: 0"+customer_phone);
                System.out.println("customer_phone: "+ customer_phone);
            } if(temp.equals("productcatagories")){
                customer_product_catagory = row.getCustomElements().getValue(tag);
                if (customer_product_catagory.charAt(customer_product_catagory.length()-1)=='s') {
                    customer_product_catagory = customer_product_catagory.substring(0, customer_product_catagory.length()-1);
                }
                jLabel_product.setText("Product: "+ customer_product_catagory);
                System.out.println("customer_product_catagory: "+ customer_product_catagory);
            } if(temp.contains("address")){
                customer_address = row.getCustomElements().getValue(tag);
                jTextArea_address.setText(customer_address);
                System.out.println("custoemr_address: "+ customer_address);
            } if(temp.equals("deliveryoptions")){
                delivery_option = row.getCustomElements().getValue(tag);
                int max_labl_size =30;
                String temp2 = "";
                if(delivery_option.length() >max_labl_size){
                    temp2 = "<html><body><p>Delivery Option:<br>" +delivery_option.substring(0, max_labl_size)+"<br>"
                            + delivery_option.substring( max_labl_size, delivery_option.length())
                            + "</p></body></html>";
                }
                
                jLabel_DeliveryOption.setText(temp2);
                System.out.println("delivery_option: "+ delivery_option);
            }
            if(temp.contains(customer_product_catagory.toLowerCase()) && temp.contains("size")){
                customer_product_size = row.getCustomElements().getValue(tag);
                int max_labl_size =40;
                String temp2 = "";
                if(customer_product_size.length() >max_labl_size){
                    temp2 = "<html><body><p>Size: <br>" +customer_product_size.substring(0, max_labl_size)+"<br>"
                            + customer_product_size.substring( max_labl_size, customer_product_size.length())
                            + "</p></body></html>";
                }
                jLabel_Size.setText(temp2);
                System.out.println("customer_product_size: "+ customer_product_size);
            }
            if(temp.contains(customer_product_catagory.toLowerCase()) && temp.contains("name")){
                customer_product_name = row.getCustomElements().getValue(tag);
                jLabel_productname.setText("Product Name: "+ customer_product_name);
                System.out.println("customer_product_name: "+ customer_product_name);
            }
            if(temp.contains(customer_product_catagory.toLowerCase()) && temp.contains("orderid")){
                orderid = row.getCustomElements().getValue(tag);
                System.out.println("customer_product_name: "+ orderid);
            }
            if(temp.contains(customer_product_catagory.toLowerCase()) && temp.contains("code")){
                customer_product_code = row.getCustomElements().getValue(tag);
                jLabel_Productcode.setText("Code: "+ customer_product_code);
                System.out.println("customer_product_code: "+ customer_product_code);
            }
            if(temp.contains(customer_product_catagory.toLowerCase()) && temp.contains("quantity")){
                customer_product_quantity = row.getCustomElements().getValue(tag);
                jTextField_quantity.setText(customer_product_quantity);
                System.out.println("customer_product_quantity: "+ customer_product_quantity);
            }
            if( temp.contains("comment")){
                jTextArea_comment.setText(row.getCustomElements().getValue(tag));
            }
            //////////////////////
            System.out.println(row.getCustomElements().getValue(tag) + "\t");
            System.out.println(customer_name+customer_phone+customer_product_name+customer_product_code+customer_product_catagory+customer_product_quantity);
            
        }
        make_description();
        if(delivery_option.toLowerCase().contains("cash on delivery")){
            jTextField_shippingCharge.setText("50");
        }else{
            jTextField_shippingCharge.setText("20");
        }
        
    }
    private static void make_description(){
        customer_product_description = customer_product_quantity+" pcs "+ customer_product_catagory + "(s)"+", Product Code(s): "+ customer_product_code + ", Size: "+ customer_product_size;
        jTextArea_Productdescription.setText(customer_product_description);
    }
    private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
        Range r1 = doc.getRange();
        
        for (int i = 0; i < r1.numSections(); ++i ) {
            Section s = r1.getSection(i);
            for (int x = 0; x < s.numParagraphs(); x++) {
                Paragraph p = s.getParagraph(x);
                for (int z = 0; z < p.numCharacterRuns(); z++) {
                    CharacterRun run = p.getCharacterRun(z);
                    String text = run.text();
                    if(text.contains(findText)) {
                        run.replaceText(findText, replaceText);
                    }
                }
            }
        }
        return doc;
    }
    private static void tempDirCreation(){
         theTempDir= mainDir+"\\TENS Posters\\Temp";
        File theTempDirFile = new File(theTempDir);
    if (!theTempDirFile.exists()) {
                //System.out.println("creating directory: " + directoryName);
                boolean result = false;
                try{
                    theTempDirFile.mkdir();
                    result = true;
                } catch(SecurityException se){
                    //handle it
                }
                if(result) {
                    System.out.println("pre-preparing tasks: TENS Posters DIR created");
                }
            }
    }
    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            out.close();
        }
    }
    private static void makeOrUpdateInvoice(){
        String filePath = "inovoice.doc";
        
        POIFSFileSystem fs = null;
        try {
            String tempInvoiceName = customer_name+"_"+customer_product_code+"_"+orderid+".doc";
            fs = new POIFSFileSystem(new FileInputStream(filePath));
            doc = new HWPFDocument(fs);
            doc = replaceText(doc, "customer_name", customer_name);
            doc = replaceText(doc, "customer_phone", "0"+customer_phone);
            doc = replaceText(doc, "customer_address", customer_address);
            System.out.println("customer_address");
            doc = replaceText(doc, "Product_name_tag", customer_product_name);
            System.out.println("Product_name_tag");
            doc = replaceText(doc, "Product_description_tag", customer_product_description);
            System.out.println("Product_description_tag");
            doc = replaceText(doc, "quantity_tag", customer_product_quantity);
            System.out.println("quantity_tag");
            doc = replaceText(doc, "price_tag", customer_product_price);
            System.out.println("price_tag");
            doc = replaceText(doc, "Sub_total_tag", customer_product_price);
            System.out.println("Sub_total_tag");
            doc = replaceText(doc, "delivery_tag", jTextField_shippingCharge.getText());
            System.out.println("delivery_tag");
            doc = replaceText(doc, "total_tag", jTextField_totalPrice.getText());
            System.out.println("total_tag");
            doc = replaceText(doc, "no_tag", "1");
            System.out.println("no_tag");
            File subDirCustome = new File(mainDir + "\\TENS Posters\\invoices");
            // if the directory does not exist, create it
            if (!subDirCustome.exists()) {
                //System.out.println("creating directory: " + directoryName);
                boolean result = false;
                
                try{
                    subDirCustome.mkdir();
                    result = true;
                } catch(SecurityException se){
                    //handle it'
                }
                if(result) {
                    System.out.println("invoices DIR created");
                }
            }
            saveWord(mainDir + "\\TENS Posters\\invoices\\"+tempInvoiceName, doc);
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
        
    }
    public static PrintService choosePrinter() {
    PrinterJob printJob = PrinterJob.getPrinterJob();
    if(printJob.printDialog()) {
        return printJob.getPrintService();          
    }
    else {
        return null;
    }
}

public static void printPDF(String fileName, PrintService printer)
        throws IOException, PrinterException {
    PrinterJob job = PrinterJob.getPrinterJob();
    job.setPrintService(printer);
    PDDocument doc = PDDocument.load(fileName);
    doc.silentPrint(job);
}
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton display_info_button;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton4;
    private static javax.swing.JButton jButton_addToECourier;
    private javax.swing.JButton jButton_makeInnovoice;
    private static javax.swing.JComboBox jComboBox2;
    private static javax.swing.JComboBox jComboBox_products;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private static javax.swing.JLabel jLabel_Contactnumber;
    private static javax.swing.JLabel jLabel_DeliveryOption;
    private static javax.swing.JLabel jLabel_Productcode;
    private static javax.swing.JLabel jLabel_Size;
    private static javax.swing.JLabel jLabel_email;
    private static javax.swing.JLabel jLabel_name;
    private static javax.swing.JLabel jLabel_product;
    private static javax.swing.JLabel jLabel_productname;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private static javax.swing.JTextArea jTextArea_Productdescription;
    private static javax.swing.JTextArea jTextArea_address;
    private static javax.swing.JTextArea jTextArea_comment;
    private javax.swing.JTextField jTextField5;
    private static javax.swing.JTextField jTextField_productPrice;
    private static javax.swing.JTextField jTextField_quantity;
    private static javax.swing.JTextField jTextField_shippingCharge;
    private static javax.swing.JTextField jTextField_totalPrice;
    private static javax.swing.JTextField row_number;
    // End of variables declaration//GEN-END:variables
}
