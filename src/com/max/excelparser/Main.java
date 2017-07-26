package com.max.excelparser;
	
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Optional;

import javax.swing.JFrame;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import com.max.excelparser.view.PanelViewController;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ProgressBar;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.text.Text;


public class Main extends Application {

	private static final int speed = 5;
    private Stage primaryStage;
    private AnchorPane rootLayout;
    private FXMLLoader mFXMLLoader;
    private PanelViewController mPanelViewController;
    private File mOldFile, mNewFile;
    private LinkedHashMap<Double, String> sRecord = new LinkedHashMap<Double, String>();
    private ProgressBar mProgressBar;
    private ImageView mCheckView_step_1, mCheckView_step_2;

    @Override
    public void start(Stage primaryStage) {
        this.primaryStage = primaryStage;
        this.primaryStage.setTitle("Excel Parser");
        initRootLayout();
        showPersonOverview();
    }

    /**
     * Initializes the root layout.
     */
    public void initRootLayout() {
        try {
            // Load root layout from fxml file.
        	mFXMLLoader= new FXMLLoader();
        	mFXMLLoader.setLocation(Main.class.getResource("view/PanelView.fxml"));
            rootLayout = (AnchorPane) mFXMLLoader.load();

            // Show the scene containing the root layout.
            Scene scene = new Scene(rootLayout);
            primaryStage.setScene(scene);
            primaryStage.show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Returns the main stage.
     * @return
     */
    public Stage getPrimaryStage() {
        return primaryStage;
    }

    public static void main(String[] args) {
        launch(args);
    }
    
    /**
     * 
     * @param title
     * @param type 0: old, 1: new
     */
	public void showFileChooser(String title, int type, Text fileName, ImageView imgView) {
	   	 FileChooser fileChooser = new FileChooser();
	   	 File initFile = new File("./");
	   	 fileChooser.setTitle(title);
	   	fileChooser.setInitialDirectory(initFile);
	   	 fileChooser.getExtensionFilters().addAll(
	   	         new ExtensionFilter("Excel Files", "*.xls"));
	   	 File selectedFile = fileChooser.showOpenDialog(primaryStage);
	   	 if (selectedFile != null) {
	 	   	fileName.setText(selectedFile.getName());
	   		 if (type == 0) {
	   			 //old file
	   			mOldFile = selectedFile;
	   			mCheckView_step_1 = imgView;
	   		 } else {
	   			 //new file
	   			 mNewFile = selectedFile;
	   			mCheckView_step_2 = imgView;
	   		 }
	   	 }
	}
    
    
    /**
     * Shows the person overview inside the root layout.
     */
    public void showPersonOverview() {
            // Give the controller access to the main app.
    	mPanelViewController = mFXMLLoader.getController();
    	mPanelViewController.setMainApp(this);
    }

	public void startCompare(ProgressBar progressbar, Text finishText) {
		mCheckView_step_1.setVisible(false);
		mCheckView_step_2.setVisible(false);
		if(finishText != null) finishText.setVisible(false);
		
		mProgressBar = progressbar;
		
        new Thread(){
            public void run() {
        		stratOldFileProcessing();
        		mCheckView_step_1.setVisible(true);
        		updateProgressBar(0);
        		startNewFileProcessing();
        		mCheckView_step_2.setVisible(true);
        		finishText.setVisible(true);
        		//showFinishDialog();
            }
        }.start();
	}

	private void showFinishDialog(){
		Platform.runLater(new Runnable() {
            @Override public void run() {
            	final Alert alert = new Alert(AlertType.CONFIRMATION); // 實體化Alert對話框物件，並直接在建構子設定對話框的訊息類型
            	alert.setTitle("完成"); //設定對話框視窗的標題列文字
            	alert.setHeaderText(""); //設定對話框視窗裡的標頭文字。若設為空字串，則表示無標頭
            	alert.setContentText("比對完成，是否離開程式"); //設定對話框的訊息文字
            	final Optional<ButtonType> opt = alert.showAndWait();
            	final ButtonType rtn = opt.get(); //可以直接用「alert.getResult()」來取代
            	//System.out.println(rtn);
            	if (rtn == ButtonType.OK) {
            	    //若使用者按下「確定」
            	    Platform.exit(); // 結束程式
            	} else if(rtn == ButtonType.CANCEL){

            	}
            }
        });
	}
	
	private void stratOldFileProcessing(){
		try {
            FileInputStream file = new FileInputStream(mOldFile);
            
            // Get the workbook instance for XLS file
            //HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // Get BugList sheet from the workbook
            HSSFSheet sheet = workbook.getSheet("BugList");
            int totalRowNum = sheet.getLastRowNum();
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            
            DataFormatter formatter = new DataFormatter();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                float CompPercent = (float)(((float)(totalRowNum- (totalRowNum - row.getRowNum())) / totalRowNum) );
                updateProgressBar(CompPercent);

                //Ignore row 0
                if(row.getRowNum() == 0) continue;
                //Get risk value, must be cell 0
                Cell riskCell = row.getCell(0);
                if(riskCell == null) continue;
                String risk = riskCell.getStringCellValue();
                if(risk == null) continue;
                //Get id value, must be cell 1
                Cell cell = row.getCell(1);
                if(cell != null) {
                    String cell_value = formatter.formatCellValue(cell);
                    if(cell_value != null && !cell_value.isEmpty()){
                        //System.out.println("num:"+row.getRowNum());
                    	double	id = Double.parseDouble(cell_value);
                    	//System.out.println("risk:"+risk);
                        sRecord.put(id, risk);
                    }
                }
                try {
                	if(row.getRowNum() % speed ==0)
                		Thread.sleep(1);
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
            //System.out.println("Key set:"+sRecord.keySet().toString());
            
            file.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void startNewFileProcessing(){
        try {
            //FileInputStream file = new FileInputStream(new File("./new.xls"));
            FileInputStream file = new FileInputStream(mNewFile);

            // Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // Get BugList sheet from the workbook
            HSSFSheet sheet = workbook.getSheet("BugList");
            int totalRowNum = sheet.getLastRowNum();
            
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            DataFormatter formatter = new DataFormatter();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                float CompPercent = (float)(((float)(totalRowNum- (totalRowNum - row.getRowNum())) / totalRowNum) );
                updateProgressBar(CompPercent);
                
                //Ignore row 0
                if(row.getRowNum() == 0) continue;
                //Get risk value
                //String risk = row.getCell(0).getStringCellValue();
                //Get id value
                Cell idCell = row.getCell(1);
                if(idCell == null) continue;
                String cell_value = formatter.formatCellValue(idCell);
                if(cell_value != null && !cell_value.isEmpty()){
                   // System.out.println("num:"+row.getRowNum());
                	double	id = Double.parseDouble(cell_value);
                    if(sRecord.containsKey(id)){                        
                        //write data to risk cell
                        Cell cell = row.getCell(0);
                        if(cell == null){
                            row.createCell(0).setCellValue(sRecord.get(id));
                        } else {
                            row.getCell(0).setCellValue(sRecord.get(id));
                        }
                    }
                }
                try {
                	if(row.getRowNum() % speed ==0)
                		Thread.sleep(1);
    			} catch (InterruptedException e) {
    				// TODO Auto-generated catch block
    				e.printStackTrace();
    			}
            }
            
            file.close();
            FileOutputStream outFile =new FileOutputStream(mNewFile);
            workbook.write(outFile);
            outFile.close();
            workbook.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void updateProgressBar(float progress){
        Platform.runLater(() -> mProgressBar.setProgress(progress));
        //System.out.println("total................."+progress);
    }
}