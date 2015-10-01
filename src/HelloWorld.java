import java.awt.BorderLayout;
import java.awt.Button;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class HelloWorld implements ActionListener{
    private static LinkedHashMap<Double, String> sRecord = new LinkedHashMap<Double, String>();
    Button mBtnNewFile, mBtnOldFile, mBtnStart;
    JLabel mLabelNewFile, mLabelOldFile, mLabelResultFile;
    File mNewSelectedFile, mOldSelectedFile;
    
    public static void main(String[] args) {
        new HelloWorld();
        //SetUpChooserUI();
        //getOldFileData();
        //getNewFileData();
        //HAHA();
    }

    private static void HAHA(){
        JFrame.setDefaultLookAndFeelDecorated(true);
        JDialog.setDefaultLookAndFeelDecorated(true);
        JFrame frame = new JFrame("JComboBox Test");
        frame.setLayout(new FlowLayout());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JButton button = new JButton("Select File");
        button.addActionListener(new ActionListener() {
          public void actionPerformed(ActionEvent ae) {
            JFileChooser fileChooser = new JFileChooser();
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
              File selectedFile = fileChooser.getSelectedFile();
              System.out.println(selectedFile.getName());
            }
          }
        });
        frame.add(button);
        frame.pack();
        frame.setVisible(true);
    }
    

    private static class dataInfo {
     String risk;
     
    }

    @Override
    public void actionPerformed(ActionEvent arg0) {
        String cmd = arg0.getActionCommand();
        if(cmd.equalsIgnoreCase("new")){
            JFileChooser fileChooser = new JFileChooser("select new");
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                mNewSelectedFile = fileChooser.getSelectedFile();
                mLabelNewFile.setText(mNewSelectedFile.getName());
            }           
        } else if(cmd.equalsIgnoreCase("old")) {
            JFileChooser fileChooser = new JFileChooser("select old");
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                mOldSelectedFile = fileChooser.getSelectedFile();
                mLabelOldFile.setText(mOldSelectedFile.getName());
            }  
        } else if(cmd.equalsIgnoreCase("start")) {
            getOldFileData();
            getNewFileData();
            mLabelResultFile.setText("DONE......................");
            //System.out.println("start");
        }
    }

    public HelloWorld(){
        JFrame mainFrame = new JFrame("Main Popup");
        mainFrame.setLayout(new GridBagLayout());
        mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //mainFrame.setSize(600, 600);

        mBtnNewFile = new Button("New File");
        mBtnOldFile = new Button("Old File");
        mBtnStart = new Button("START");

        mBtnNewFile.addActionListener(this);
        mBtnOldFile.addActionListener(this);
        mBtnStart.addActionListener(this);

        mBtnNewFile.setActionCommand("new");
        mBtnOldFile.setActionCommand("old");
        mBtnStart.setActionCommand("start");
        
        GridBagConstraints c0 = new GridBagConstraints();
        c0.gridx = 0;
        c0.gridy = 0;
        c0.gridwidth = 2;
        c0.gridheight = 1;
        c0.weightx = 0;
        c0.weighty = 0;
        c0.fill = GridBagConstraints.NONE;
        c0.anchor = GridBagConstraints.EAST;
        mainFrame.add(mBtnNewFile, c0);
        
        GridBagConstraints c1 = new GridBagConstraints();
        c1.gridx = 0;
        c1.gridy = 1;
        c1.gridwidth = 2;
        c1.gridheight = 1;
        c1.weightx = 0;
        c1.weighty = 0;
        c1.fill = GridBagConstraints.NONE;
        c1.anchor = GridBagConstraints.WEST;
        mainFrame.add(mBtnOldFile, c1);
        
        mLabelNewFile = new JLabel();
        GridBagConstraints c3 = new GridBagConstraints();
        c3.gridx = 2;
        c3.gridy = 0;
        c3.gridwidth = 6;
        c3.gridheight = 1;
        c3.weightx = 0;
        c3.weighty = 0;
        c3.fill = GridBagConstraints.BOTH;
        c3.anchor = GridBagConstraints.WEST;
        mainFrame.add(mLabelNewFile, c3);
         
        mLabelOldFile = new JLabel();
        GridBagConstraints c4 = new GridBagConstraints();
        c4.gridx = 2;
        c4.gridy = 1;
        c4.gridwidth = 6;
        c4.gridheight = 1;
        c4.weightx = 0;
        c4.weighty = 0;
        c4.fill = GridBagConstraints.BOTH;
        c4.anchor = GridBagConstraints.WEST;
        mainFrame.add(mLabelOldFile, c4);

        GridBagConstraints c5 = new GridBagConstraints();
        c5.gridx = 0;
        c5.gridy = 14;
        c5.gridwidth = 16;
        c5.gridheight = 2;
        c5.weightx = 0;
        c5.weighty = 0;
        c5.fill = GridBagConstraints.HORIZONTAL;
        c5.anchor = GridBagConstraints.SOUTHWEST;
        mainFrame.add(mBtnStart, c5);
        
        mLabelResultFile = new JLabel("OHOH");
        GridBagConstraints c6 = new GridBagConstraints();
        c6.gridx = 0;
        c6.gridy = -1;
        c6.gridwidth = 0;
        c6.gridheight = 0;
        c6.weightx = 0;
        c6.weighty = 0;
        c6.fill = GridBagConstraints.NONE;
        c6.anchor = GridBagConstraints.EAST;
        mainFrame.add(mLabelResultFile, c6);
        
        // make the frame half the height and width
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int height = screenSize.height;
        int width = screenSize.width;
        mainFrame.setSize(width/2, height/2);
        //Container contentPane = frame.getContentPane();
        
        
        
        
        mainFrame.setLocationRelativeTo(null);
        //mainFrame.pack();
        mainFrame.setVisible(true);
    }

    private static void SetUpChooserUI(){
        JFrame frame = new JFrame("JFileChooser Popup");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JFrame.setDefaultLookAndFeelDecorated(true);
        Container contentPane = frame.getContentPane();

        final JLabel directoryLabel = new JLabel(" ");
        directoryLabel.setFont(new Font("Serif", Font.BOLD | Font.ITALIC, 36));
        contentPane.add(directoryLabel, BorderLayout.NORTH);

        final JLabel filenameLabel = new JLabel(" ");
        filenameLabel.setFont(new Font("Serif", Font.BOLD | Font.ITALIC, 36));
        contentPane.add(filenameLabel, BorderLayout.SOUTH);

        JFileChooser fileChooser = new JFileChooser();
        //fileChooser.setControlButtonsAreShown(false);
        contentPane.add(fileChooser, BorderLayout.CENTER);

        ActionListener actionListener = new ActionListener() {
            public void actionPerformed(ActionEvent actionEvent) {
              JFileChooser theFileChooser = (JFileChooser) actionEvent
                  .getSource();
              String command = actionEvent.getActionCommand();
              if (command.equals(JFileChooser.APPROVE_SELECTION)) {
                File selectedFile = theFileChooser.getSelectedFile();
                directoryLabel.setText(selectedFile.getParent());
                filenameLabel.setText(selectedFile.getName());
              } else if (command.equals(JFileChooser.CANCEL_SELECTION)) {
                directoryLabel.setText(" ");
                filenameLabel.setText(" ");
              }
            }
          };
          fileChooser.addActionListener(actionListener);

          frame.pack();
          frame.setVisible(true);
    }

    private void getOldFileData(){
        try {
            
            //FileInputStream file = new FileInputStream(new File("./old.xls"));
            FileInputStream file = new FileInputStream(mOldSelectedFile);
            
            // Get the workbook instance for XLS file
            //HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // Get BugList sheet from the workbook
            HSSFSheet sheet = workbook.getSheet("BugList");
            int totalRowNum = sheet.getLastRowNum();
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                int CompPercent = (int)(((float)(totalRowNum- (totalRowNum - row.getRowNum())) / totalRowNum) * 100);
                String complete = String.valueOf(CompPercent + "%");
                System.out.println("total................."+complete);
                mLabelResultFile.setText(complete);
                //Ignore row 0
                if(row.getRowNum() == 0) continue;
                //Get risk value, must be cell 0
                Cell riskCell = row.getCell(0);
                if(riskCell == null) continue;
                String risk = riskCell.getStringCellValue();
                if(risk == null) continue;
                //Get id value, must be cell 1
                if(row.getCell(1) != null){
                    double id = row.getCell(1).getNumericCellValue();
                    sRecord.put(id, risk);
                }
            }
            System.out.println("Key set:"+sRecord.keySet().toString());
            
            file.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
    }

    private void getNewFileData(){
        try {
            //FileInputStream file = new FileInputStream(new File("./new.xls"));
            FileInputStream file = new FileInputStream(mNewSelectedFile);

            // Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // Get BugList sheet from the workbook
            HSSFSheet sheet = workbook.getSheet("BugList");
            
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //Ignore row 0
                if(row.getRowNum() == 0) continue;
                //Get risk value
                //String risk = row.getCell(0).getStringCellValue();
                //Get id value
                Cell idCell = row.getCell(1);
                if(idCell == null) continue;
                double id = idCell.getNumericCellValue();
                if(sRecord.containsKey(id)){
                    //System.out.println("found:"+sRecord.get(id));
                    
                    //write data to risk cell
                    Cell cell = row.getCell(0);
                    if(cell == null){
                        row.createCell(0).setCellValue(sRecord.get(id));
                    } else {
                        row.getCell(0).setCellValue(sRecord.get(id));
                    }
                }
                //sRecord.put(id, risk);
            }
            file.close();            
            //FileOutputStream outFile =new FileOutputStream(new File("./new.xls"));

            FileOutputStream outFile =new FileOutputStream(mNewSelectedFile);
            workbook.write(outFile);
            outFile.close();
            workbook.close();
            JFrame frame = new JFrame("JFileChooser Popup");
            //JOptionPane.showMessageDialog(frame, "FINSH..............");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
    }
}