package com.max.excelparser.view;

import java.io.File;

import com.max.excelparser.Main;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextField;
import javafx.scene.image.ImageView;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;

public class PanelViewController {
    @FXML
    private Button mOkBtn;
    
    @FXML
    private Button mCancelBtn;
    
    @FXML
    private Button mOldFileBtn;
    
    @FXML
    private Button mNewFileBtn;
    
    @FXML
    private Text mNewFileName;
    
    @FXML
    private Text mOldFileName;

    @FXML
    private ProgressBar mProgressBar;
    
    @FXML
    private ImageView mCheckView_step_1;
    @FXML
    private ImageView mCheckView_step_2;
    
    @FXML
    private Text mFinishText;;
    
    // Reference to the main application.
    private Main mainApp;

    /**
     * The constructor.
     * The constructor is called before the initialize() method.
     */
    public PanelViewController() {
    }

    /**
     * Initializes the controller class. This method is automatically called
     * after the fxml file has been loaded.
     */
    @FXML
    private void initialize() {
    }

    @FXML
    private void onOldFileBtnClickHandle(){
    	mainApp.showFileChooser("請選擇 舊 的檔案", 0, mOldFileName, mCheckView_step_1);
    }

    @FXML
    private void onNewFileBtnClickHandle(){
    	mainApp.showFileChooser("請選擇 新 的檔案", 1, mNewFileName, mCheckView_step_2);
    }

    @FXML
    private void onOkBtnClickHandle(){
    	mainApp.startCompare(mProgressBar, mFinishText);
    }

    @FXML
    private void onCancelBtnClickHandle(){
    	//System.out.println("Cancel click");
        Stage stage = (Stage) mCancelBtn.getScene().getWindow();
        stage.close();
    }


    /**
     * Is called by the main application to give a reference back to itself.
     * 
     * @param mainApp
     */
    public void setMainApp(Main mainApp) {
        this.mainApp = mainApp;
        mCheckView_step_1.setVisible(false);
        mCheckView_step_2.setVisible(false);
        mFinishText.setVisible(false);
    }

}