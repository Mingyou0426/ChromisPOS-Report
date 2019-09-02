/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package salesreport;

import java.awt.Cursor;
import java.awt.Font;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JFormattedTextField.AbstractFormatter;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPasswordField;
import javax.swing.JRootPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.jdatepicker.impl.JDatePanelImpl;
import org.jdatepicker.impl.JDatePickerImpl;
import org.jdatepicker.impl.UtilDateModel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.view.JasperViewer;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.util.JRLoader;

public class SalesReport {
    JFrame window; // This is Main Window
    JFrame dbsetwnd; // This is Main Window

    JTextField txtExcelPath;
    JButton btnOutFile;
    JLabel lblExcelPath;

    JLabel lblComment;

    JLabel lblFromDate;
    JDatePanelImpl fromdatePanel;
    JDatePickerImpl fromdatePicker;

    JDatePanelImpl todatePanel;
    JDatePickerImpl todatePicker;
    JLabel lblToDate;


    UtilDateModel frommodel;
    UtilDateModel tomodel;
    Properties p;

    String excelFileName = "";

    JButton btnDBSettings;

    JTextField txtDBUser;
    JLabel lblDBUser;
    JPasswordField txtDBPass;
    JLabel lblDBPass;

    JTextField txtDBHost;
    JLabel lblDBHost;
    JTextField txtDBName;
    JLabel lblDBName;

    JButton btnOK;
    JButton btnCancel;

    JCheckBox chkShow;

    public class DateLabelFormatter extends AbstractFormatter {
        private String datePattern = "yyyy-MM-dd";
        private SimpleDateFormat dateFormatter = new SimpleDateFormat(datePattern);

        @Override
        public Object stringToValue(String text) throws ParseException {
            return dateFormatter.parseObject(text);
        }

        @Override
        public String valueToString(Object value) throws ParseException {
            if (value != null) {
                Calendar cal = (Calendar) value;
                return dateFormatter.format(cal.getTime());
            }

            return "";
        }

    }

    SalesReport () {
        window = new JFrame("Sales Report");
        window.setSize(600, 180); // Height And Width Of Window
        window.setLocationRelativeTo(null); // Move Window To Center

        lblComment = new JLabel("Choose date range for report.");
        lblComment.setBounds(100, 20, 400, 30);
        lblComment.setFont(new Font("Comic Sans MS", Font.PLAIN, 25));
        window.add(lblComment);

        btnDBSettings = new JButton();
        btnDBSettings.setBounds(510, 20, 30, 30);
        btnDBSettings.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnDBSettings.setToolTipText("Database Access Information Settings");
        btnDBSettings.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                dbsetwnd.setVisible(true);
            }
        });
        try {
            Image img = ImageIO.read(getClass().getResource("resources/settings.png"));
            Image newimg = img.getScaledInstance( 30, 30,  java.awt.Image.SCALE_SMOOTH ) ;
            btnDBSettings.setIcon(new ImageIcon(newimg));
        } catch (Exception ex) {
            System.out.println(ex);
        }
        window.add(btnDBSettings);

        lblExcelPath = new JLabel("Excel File Name");
        lblExcelPath.setBounds(30, 100, 120 , 30);
        lblExcelPath.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        window.add(lblExcelPath);

        txtExcelPath = new JTextField();
        txtExcelPath.setBounds(160, 100, 400 , 30);
        txtExcelPath.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        txtExcelPath.setEnabled(false);
        window.add(txtExcelPath);

        frommodel = new UtilDateModel();
        tomodel = new UtilDateModel();
        p = new Properties();
        p.put("text.today", "Today");
        p.put("text.month", "Month");
        p.put("text.year", "Year");

        lblFromDate = new JLabel("From Date");
        lblFromDate.setBounds(20, 60, 80 , 30);
        lblFromDate.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        window.add(lblFromDate);

        fromdatePanel = new JDatePanelImpl(frommodel, p);
        fromdatePicker = new JDatePickerImpl(fromdatePanel, new DateLabelFormatter());
        fromdatePicker.setBounds(110, 60, 150, 30);
        fromdatePicker.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                excelFileName = "SalesReport_"+fromdatePicker.getJFormattedTextField().getText()+"_"+
                        todatePicker.getJFormattedTextField().getText()+".xlsx";
                txtExcelPath.setText(excelFileName);
            }
        });
        window.add(fromdatePicker);

        lblToDate = new JLabel("To Date");
        lblToDate.setBounds(270, 60, 60 , 30);
        lblToDate.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        lblToDate.setEnabled(true);
        window.add(lblToDate);

        todatePanel = new JDatePanelImpl(tomodel, p);
        todatePicker = new JDatePickerImpl(todatePanel, new DateLabelFormatter());
        todatePicker.setBounds(340, 60, 150, 30);
        todatePicker.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                excelFileName = "SalesReport_"+fromdatePicker.getJFormattedTextField().getText()+"_"+
                        todatePicker.getJFormattedTextField().getText()+".xlsx";
                txtExcelPath.setText(excelFileName);
            }
        });
        window.add(todatePicker);

        btnOutFile = new JButton();
        btnOutFile.setText("Report");
        btnOutFile.setBounds(500, 60, 80, 30);
        btnOutFile.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnOutFile.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                if ( fromdatePicker.getJFormattedTextField().getText().trim().equals("") ) {
                    JOptionPane.showMessageDialog(null, "Please Input From Date!", "Error Message", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }

                if ( !todatePicker.getJFormattedTextField().getText().trim().equals("") &&
                        todatePicker.getJFormattedTextField().getText().trim().compareTo(fromdatePicker.getJFormattedTextField().getText().trim())<0 ) {
                    JOptionPane.showMessageDialog(null, "To Date is before than From Date!", "Error Message", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }

                Date fromdate = frommodel.getValue();
                Date todate;

                if ( todatePicker.getJFormattedTextField().getText().trim().equals("") ) {
                    todate = frommodel.getValue();
                } else {
                    todate = tomodel.getValue();
                }
                
                Connection conn = null;
                Statement stmt = null;
                ResultSet rs = null;
                Statement stmt1 = null;
                ResultSet rs1 = null;
                List<SalesReportModel> modelList = new ArrayList<SalesReportModel>();

                String format = "dd-MMM-yy";
                DateFormat df = new SimpleDateFormat(format);
                try {
                    String url = "jdbc:mysql://"+txtDBHost.getText().trim()+":3306/"+txtDBName.getText().trim()+"?" + "useSSL=false";
                    String user      = txtDBUser.getText().trim();
                    String password  = String.valueOf(txtDBPass.getPassword());
                    Class.forName("com.mysql.jdbc.Driver");
                    conn = DriverManager.getConnection(url, user, password);

                    stmt = conn.createStatement();
                    stmt1 = conn.createStatement();

                    String excelFileName = txtExcelPath.getText();
                    String sheetName = "Sheet1";//name of sheet
                    XSSFWorkbook wb = new XSSFWorkbook();
                    XSSFSheet sheet = wb.createSheet(sheetName);

                    for (int i=0; i<5; i++){
                        sheet.autoSizeColumn(i);
                    }

                    XSSFRow row = sheet.createRow(0);
                    XSSFCell cell = row.createCell(0);
                    cell.setCellValue("Date");
                    cell = row.createCell(1);
                    cell.setCellValue("Quantity sold");
                    cell = row.createCell(2);
                    cell.setCellValue("Cost Value");
                    cell = row.createCell(3);
                    cell.setCellValue("Sales Value");
                    cell = row.createCell(4);
                    cell.setCellValue("Profit");

                    Calendar cal = Calendar.getInstance();
                    cal.setTime(fromdate);
                    String sql = "";
                    int quantitysold;
                    int costvalue;
                    int salesvalue;
                    int profit;
                    int tquantitysold = 0;
                    int tcostvalue = 0;
                    int tsalesvalue = 0;
                    int tprofit = 0;
                    CellStyle cellStyle;
                    CreationHelper createHelper = wb.getCreationHelper();
                    FormulaEvaluator fev = createHelper.createFormulaEvaluator();
                    int rownum = 1;
                    do {
                        row = sheet.createRow(rownum);
                        cell = row.createCell(0);
                        cellStyle = wb.createCellStyle();
                        cellStyle.setDataFormat(
                            createHelper.createDataFormat().getFormat("d-mmm-yy"));
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(cal.getTime());
                        
                        quantitysold = 0;
                        costvalue = 0;
                        salesvalue = 0;
                        profit = 0;
                        String tmp = String.format("%04d-%02d-%02d", cal.get(Calendar.YEAR), cal.get(Calendar.MONTH)+1, cal.get(Calendar.DATE));
                        sql = "select B.UNITS, C.PRICEBUY, B.PRICE from receipts A, ticketlines B, products C, payments D where A.DATENEW LIKE '"+tmp+
                                "%' and A.ID=B.TICKET"+
                                " and B.PRODUCT=C.ID and D.RECEIPT=A.ID and D.PAYMENT='cash'";
                        rs = stmt.executeQuery(sql);
                        while ( rs.next() ) {
                            quantitysold = quantitysold+rs.getInt("UNITS");
                            tquantitysold = tquantitysold+rs.getInt("UNITS");

                            costvalue = costvalue+rs.getInt("PRICEBUY")*rs.getInt("UNITS");
                            tcostvalue = tcostvalue+rs.getInt("PRICEBUY")*rs.getInt("UNITS");

                            salesvalue = salesvalue+rs.getInt("PRICE")*rs.getInt("UNITS");
                            tsalesvalue = tsalesvalue+rs.getInt("PRICE")*rs.getInt("UNITS");

                            profit = profit+(rs.getInt("PRICE")-rs.getInt("PRICEBUY"))*rs.getInt("UNITS");
                            tprofit = tprofit+(rs.getInt("PRICE")-rs.getInt("PRICEBUY"))*rs.getInt("UNITS");
                        }

                        modelList.add(new SalesReportModel(df.format(cal.getTime()),
                                String.valueOf(quantitysold),
                                String.valueOf(costvalue),
                                String.valueOf(salesvalue),
                                String.valueOf(profit)));

                        cell = row.createCell(1);
                        cell.setCellFormula("VALUE(" + String.valueOf(quantitysold) + ")");
                        fev.evaluateInCell(cell);                        

                        cell = row.createCell(2);
                        cell.setCellFormula("VALUE(" + String.valueOf(costvalue) + ")");
                        fev.evaluateInCell(cell);

                        cell = row.createCell(3);
                        cell.setCellFormula("VALUE(" + String.valueOf(salesvalue) + ")");
                        fev.evaluateInCell(cell);

                        cell = row.createCell(4);
                        // cell.setCellFormula("VALUE(" + String.valueOf(profit) + ")");
                        cell.setCellFormula("D"+String.valueOf(rownum+1)+"-C"+String.valueOf(rownum+1));
                        fev.evaluateInCell(cell);

                        cal.add(Calendar.DAY_OF_MONTH, 1);
                        rownum++;
                    } while ( !cal.getTime().after(todate) );

                    rownum++;

                    XSSFFont font= wb.createFont();
                    font.setBold(true);

                    row = sheet.createRow(rownum);
                    cell = row.createCell(0);
                    cellStyle = wb.createCellStyle();
                    cellStyle.setFont(font);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue("Grand Total");

                    cell = row.createCell(1);
                    cell.setCellStyle(cellStyle);
                    cell.setCellFormula("SUM(B2:B"+String.valueOf(rownum-1)+")");
                    fev.evaluateInCell(cell);

                    cellStyle = wb.createCellStyle();
                    cellStyle.setFont(font);
                    cellStyle.setDataFormat(
                            createHelper.createDataFormat().getFormat("_ [$₦-470] * #,##0.00_ ;_ [$₦-470] * -#,##0.00_ ;_ [$₦-470] * \"-\"??_ ;_ @_ "));
                    cell = row.createCell(2);
                    cell.setCellStyle(cellStyle);
                    cell.setCellFormula("SUM(C2:C"+String.valueOf(rownum-1)+")");
                    fev.evaluateInCell(cell);

                    cell = row.createCell(3);
                    cell.setCellStyle(cellStyle);
                    cell.setCellFormula("SUM(D2:D"+String.valueOf(rownum-1)+")");
                    fev.evaluateInCell(cell);

                    cell = row.createCell(4);
                    cell.setCellStyle(cellStyle);
                    cell.setCellFormula("SUM(E2:E"+String.valueOf(rownum-1)+")");
                    fev.evaluateInCell(cell);

                    for (int i=0; i<5; i++){
                        sheet.autoSizeColumn(i);
                    }
                    // wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
                    FileOutputStream fileOut = new FileOutputStream(excelFileName);
                    wb.write(fileOut);
                    fileOut.flush();
                    fileOut.close();

                    if(conn != null)
                        conn.close();
                    if(stmt != null)
                        stmt.close();
                    if(rs != null)
                        rs.close();
                    if(stmt1 != null)
                        stmt1.close();
                    if(rs1 != null)
                        rs1.close();

                    modelList.add(new SalesReportModel("Grand Total",
                            String.valueOf(tquantitysold),
                            "₦   "+String.format("%.2f", Float.valueOf(tcostvalue)),
                            "₦   "+String.format("%.2f", Float.valueOf(tsalesvalue)),
                            "₦   "+String.format("%.2f", Float.valueOf(tprofit))));

                    String sourceFileName = "SalesReport.jrxml";
                    JasperReport jasperReport = null;
                    jasperReport = JasperCompileManager.compileReport(sourceFileName);
                    JRBeanCollectionDataSource dataSource = new JRBeanCollectionDataSource(modelList);
                    Map<String,Object> params = new HashMap<String,Object>();
                    JasperPrint jasperPrint = JasperFillManager.fillReport(jasperReport, params, dataSource);
                    String path = "SalesReport_"+fromdatePicker.getJFormattedTextField().getText()+"_"+
                        todatePicker.getJFormattedTextField().getText()+".pdf";
                    JasperExportManager.exportReportToPdfFile(jasperPrint,path);

                    JasperViewer jv = new JasperViewer(jasperPrint);
                    JDialog viewer = new JDialog(window, "SalesReport_"+fromdatePicker.getJFormattedTextField().getText()+"_"+
                        todatePicker.getJFormattedTextField().getText(), true);
                    viewer.setBounds(jv.getBounds());
                    viewer.getContentPane().add(jv.getContentPane());
                    viewer.setResizable(true);
                    viewer.setIconImage(jv.getIconImage());
                    viewer.setVisible(true);

                    JOptionPane.showMessageDialog(null, "Successfully Done!", "OK Message", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex1) {
                    JOptionPane.showMessageDialog(null, "Error occured.", "Error Message", JOptionPane.INFORMATION_MESSAGE);
                    System.out.println(ex1.getMessage());

                    PrintWriter writer;
                    try {
                        writer = new PrintWriter("log.txt", "UTF-8");
                        writer.println(ex1.getMessage());
                        writer.close();
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(SalesReport.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (UnsupportedEncodingException ex) {
                        Logger.getLogger(SalesReport.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    return;
                } finally {
                    try {
                        if(conn != null)
                            conn.close();
                        if(stmt != null)
                            stmt.close();
                        if(rs != null)
                            rs.close();
                        if(stmt1 != null)
                            stmt1.close();
                        if(rs1 != null)
                            rs1.close();
                    } catch(SQLException ex) {
                        System.out.println(ex.getMessage());
                    }
                }
            }
        });
        window.add(btnOutFile);
        
        window.setLayout(null);
        window.setResizable(false);
        window.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // If Click into The Red Button => End The Processus
        window.setVisible(true);

        dbsetwnd = new JFrame("DB Settings");
        dbsetwnd.setSize(240,280);
        dbsetwnd.setLocationRelativeTo(null);
        dbsetwnd.setResizable(false);
        dbsetwnd.setLayout(null);
        dbsetwnd.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);

        lblDBUser = new JLabel("DB User");
        lblDBUser.setBounds(30, 110, 60, 30);
        lblDBUser.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(lblDBUser);

        txtDBUser = new JTextField();
        txtDBUser.setBounds(100, 110, 100 , 30);
        txtDBUser.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(txtDBUser);

        lblDBPass = new JLabel("DB Pass");
        lblDBPass.setBounds(30, 150, 60, 30);
        lblDBPass.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(lblDBPass);

        txtDBPass = new JPasswordField();
        txtDBPass.setBounds(100, 150, 100 , 30);
        txtDBPass.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(txtDBPass);

        lblDBHost = new JLabel("DB Host");
        lblDBHost.setBounds(30, 30, 60, 30);
        lblDBHost.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(lblDBHost);

        txtDBHost = new JTextField();
        txtDBHost.setBounds(100, 30, 100 , 30);
        txtDBHost.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(txtDBHost);

        lblDBName = new JLabel("DB Name");
        lblDBName.setBounds(30, 70, 70, 30);
        lblDBName.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(lblDBName);

        txtDBName = new JTextField();
        txtDBName.setBounds(100, 70, 100 , 30);
        txtDBName.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        dbsetwnd.add(txtDBName);

        chkShow = new JCheckBox("");
        chkShow.setBounds(200, 150, 30, 30);
        chkShow.setToolTipText("Show Password");
        chkShow.setFont(new Font("Comic Sans MS", Font.PLAIN, 15));
        chkShow.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                if ( chkShow.isSelected() ) {
                    txtDBPass.setEchoChar((char)0);
                } else {
                    txtDBPass.setEchoChar('*');
                }
            }
        });
        dbsetwnd.add(chkShow);

        btnOK = new JButton();
        btnOK.setText("OK");
        btnOK.setBounds(30, 190, 80, 30);
        btnOK.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnOK.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                txtDBUser.setText(txtDBUser.getText().trim());
                if ( txtDBUser.getText().equals("") ) {
                    JOptionPane.showMessageDialog(null, "Please Input DB User!", "Error Message", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }
                PrintWriter writer;
                try {
                    writer = new PrintWriter("db.ini", "UTF-8");
                    writer.println(txtDBHost.getText());
                    writer.println(txtDBName.getText());
                    writer.println(txtDBUser.getText());
                    writer.println(String.valueOf(txtDBPass.getPassword()));
                    writer.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(SalesReport.class.getName()).log(Level.SEVERE, null, ex);
                } catch (UnsupportedEncodingException ex) {
                    Logger.getLogger(SalesReport.class.getName()).log(Level.SEVERE, null, ex);
                }

                dbsetwnd.setVisible(false);
            }
        });
        JRootPane rootPane = SwingUtilities.getRootPane(dbsetwnd);
        rootPane.setDefaultButton(btnOK);
        dbsetwnd.getContentPane().add(btnOK);

        btnCancel = new JButton();
        btnCancel.setText("Cancel");
        btnCancel.setBounds(120, 190, 80, 30);
        btnCancel.setCursor(new Cursor(Cursor.HAND_CURSOR));
        btnCancel.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){
                dbsetwnd.setVisible(false);
            }
        });
        dbsetwnd.add(btnCancel);

        try {
            BufferedReader br = new BufferedReader(new FileReader("db.ini"));
            try {
                String dbhost = br.readLine();
                String dbname = br.readLine();
                String dbuser = br.readLine();
                String dbpass = br.readLine();

                txtDBHost.setText(dbhost);
                txtDBName.setText(dbname);
                txtDBUser.setText(dbuser);
                txtDBPass.setText(dbpass);
            } finally {
                br.close();
            }
        } catch (Exception ex) {
            Logger.getLogger(SalesReport.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void main(String[] args) {
        new SalesReport();
    }
}
