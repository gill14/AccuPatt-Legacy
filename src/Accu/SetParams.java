package Accu;

import com.oceanoptics.omnidriver.api.wrapper.Wrapper;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import jssc.SerialPort;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.Optional;

/**
 * Created by gill14 on 2/4/2016.
 */
public class SetParams {

    public static long timeStart;
    public static long timeStop;
    public static double stringCalFactor;

    public static void settings(String params, int excitationWavelength, int targetWavelength, int boxcarWidth, int integrationTime, boolean metric, double flightlineLength, double sampleLength, double stringTimeFactor, boolean isFlipPattern, boolean useLogo, String logo, boolean invertPH, SerialPort serialPort, int sFIdentifier, double spreadA, double spreadB, double spreadC, int threshold, double crop, int spacing){
        Stage window = new Stage();
        window.initModality(Modality.APPLICATION_MODAL);
        window.setTitle("Settings");
        //window.setMinWidth(300);

        //Divide things up!
        HBox hBoxTop = new HBox(5);
        VBox vBoxLeft = new VBox(10);
        Separator separator0 = new Separator(Orientation.VERTICAL);

        stringCalFactor = stringTimeFactor;

        //Spectrometer Settings

        HBox hBoxSpectLabel = new HBox(5);
        Text spectLabel = new Text("Spectrometer");
        spectLabel.setFont(Font.font(null, FontWeight.BOLD, 24));
        spectLabel.setFill(Color.WHITESMOKE);
        hBoxSpectLabel.getChildren().add(spectLabel);
        hBoxSpectLabel.setStyle("-fx-background-color: gray");

        HBox hB0 = new HBox(10);
        Label lbDye = new Label("Tracer: ");
        Region regDye = new Region();
        hB0.setHgrow(regDye, Priority.ALWAYS);
        ToggleGroup tGG0 = new ToggleGroup();
        RadioButton rB000 = new RadioButton("Rhodamine");
        rB000.setToggleGroup(tGG0);
        RadioButton rB001 = new RadioButton("Pyranine");
        rB001.setToggleGroup(tGG0);
        RadioButton rB002 = new RadioButton("Custom");
        rB002.setToggleGroup(tGG0);

        hB0.getChildren().addAll(lbDye,regDye,rB000,rB001,rB002);

        HBox hBExWv = new HBox(10);
        Label lbExWv = new Label("Excitation Wavelength (nm): ");
        Region regExWv = new Region();
        hBExWv.setHgrow(regExWv, Priority.ALWAYS);
        TextField tFExWv = new TextField(Integer.toString(excitationWavelength));
        hBExWv.getChildren().addAll(lbExWv, regExWv, tFExWv);

        HBox hBEmWv = new HBox(10);
        Label lB0 = new Label("Emission Wavelength (nm): ");
        Region reg0 = new Region();
        hBEmWv.setHgrow(reg0, Priority.ALWAYS);
        TextField tFTargetWavelength = new TextField(Integer.toString(targetWavelength));
        hBEmWv.getChildren().addAll(lB0, reg0, tFTargetWavelength);

        HBox hB1 = new HBox(10);
        Label lB1 = new Label("Boxcar Width: ");
        Region reg1 = new Region();
        hB1.setHgrow(reg1, Priority.ALWAYS);
        ToggleGroup tGG = new ToggleGroup();
        RadioButton rB00 = new RadioButton("9");
        rB00.setToggleGroup(tGG);
        if (boxcarWidth==9) {
            rB00.setSelected(true);
        } else {
            rB00.setSelected(false);
        }
        RadioButton rB01 = new RadioButton("21");
        rB01.setToggleGroup(tGG);
        if (boxcarWidth==21) {
            rB01.setSelected(true);
        } else {
            rB01.setSelected(false);
        }
        RadioButton rB02 = new RadioButton("41");
        rB02.setToggleGroup(tGG);
        if (boxcarWidth==41) {
            rB02.setSelected(true);
        } else {
            rB02.setSelected(false);
        }
        RadioButton rB03 = new RadioButton("61");
        rB03.setToggleGroup(tGG);
        if (boxcarWidth==61) {
            rB03.setSelected(true);
        } else {
            rB03.setSelected(false);
        }
        hB1.getChildren().addAll(lB1,reg1,rB00,rB01,rB02,rB03);

        HBox hB2 = new HBox(10);
        Label lB2 = new Label("Integration Time ("+"\u00B5"+"s): ");
        Region reg2 = new Region();
        hB2.setHgrow(reg2, Priority.ALWAYS);
        TextField tFIntTime = new TextField(Integer.toString(integrationTime));
        hB2.getChildren().addAll(lB2,reg2,tFIntTime);

        //Set Dye Buttons
        if(targetWavelength==575){
            rB000.setSelected(true);
        } else if(targetWavelength==495){
            rB001.setSelected(true);
        } else {
            rB002.setSelected(true);
        }

        rB000.setOnMouseClicked(e -> {
            tFExWv.setText("525");
            tFTargetWavelength.setText("575");
        });
        rB001.setOnMouseClicked(e -> {
            tFExWv.setText("425");
            tFTargetWavelength.setText("495");
        });

        HBox hBSpectTest = new HBox(10);
        Region regST1 = new Region();
        hBSpectTest.setHgrow(regST1, Priority.ALWAYS);
        Button btnSpectTest = new Button("Test Spectrometer");
        btnSpectTest.setOnMouseClicked(e->{
            new ViewSpectrum(integrationTime, boxcarWidth, excitationWavelength, targetWavelength);
        });
        Region regST2 = new Region();
        hBSpectTest.setHgrow(regST2, Priority.ALWAYS);
        hBSpectTest.getChildren().addAll(regST1, btnSpectTest, regST2);

        //******************************************************************************************

        //Label stringDriveLabel = new Label("String Drive Settings");
        HBox hBoxStringDriveLabel = new HBox(5);
        Text stringDriveLabel = new Text("String Drive");
        stringDriveLabel.setFont(Font.font(null, FontWeight.BOLD, 24));
        stringDriveLabel.setFill(Color.WHITESMOKE);
        hBoxStringDriveLabel.getChildren().addAll(stringDriveLabel);
        hBoxStringDriveLabel.setStyle("-fx-background-color: gray");

        HBox hBMetric = new HBox(10);
        Label lBMetric = new Label("Units");
        Region regMetric = new Region();
        hBMetric.setHgrow(regMetric, Priority.ALWAYS);
        ToggleGroup tGM = new ToggleGroup();
        RadioButton rBM0 = new RadioButton("US");
        rBM0.setToggleGroup(tGM);
        if (!metric) {
            rBM0.setSelected(true);
        } else {
            rBM0.setSelected(false);
        }
        RadioButton rBM1 = new RadioButton("Metric");
        rBM1.setToggleGroup(tGM);
        if (metric) {
            rBM1.setSelected(true);
        } else {
            rBM1.setSelected(false);
        }
        hBMetric.getChildren().addAll(lBMetric, regMetric, rBM0, rBM1);

        HBox hB4 = new HBox(10);
        Label lB4;
        Region reg4;
        TextField tFSampleLength;
        if (!metric) {
            lB4 = new Label("Sample Length (ft)");
            reg4 = new Region();
            hB4.setHgrow(reg4, Priority.ALWAYS);
            tFSampleLength = new TextField(Double.toString(sampleLength));
        } else {
            lB4 = new Label("Sample Length (m)");
            reg4 = new Region();
            hB4.setHgrow(reg4, Priority.ALWAYS);
            tFSampleLength = new TextField(Double.toString(sampleLength));
        }
        hB4.getChildren().addAll(lB4,reg4,tFSampleLength);

        HBox hB3 = new HBox(10);
        Label lB3;
        Region reg3;
        TextField tFFlightlineLength;
        if (!metric) {
            lB3 = new Label("Flightline Length (ft)");
            reg3 = new Region();
            hB3.setHgrow(reg3, Priority.ALWAYS);
            tFFlightlineLength = new TextField(Integer.toString((int)flightlineLength));
        } else {
            lB3 = new Label("Flightline Length (m)");
            reg3 = new Region();
            hB3.setHgrow(reg3, Priority.ALWAYS);
            tFFlightlineLength = new TextField(Double.toString(flightlineLength));
        }
        hB3.getChildren().addAll(lB3,reg3,tFFlightlineLength);



        HBox hB5 = new HBox(10);
        Label lB5;
        TextField tFStringTimeFactor;
        if (!metric) {
            lB5 = new Label("Calibrate String Drive (ft/s): ");
            tFStringTimeFactor = new TextField(Double.toString(stringTimeFactor));
        } else {
            lB5 = new Label("Calibrate String Drive (m/s): ");
            tFStringTimeFactor = new TextField(Double.toString(stringTimeFactor));
        }
        tFStringTimeFactor.setPrefWidth(40);
        Region reg5 = new Region();
        hB5.setHgrow(reg5, Priority.ALWAYS);
        Button buttonStart = new Button("Start");
        buttonStart.setStyle("-fx-background-color: #009900");
        Button buttonStop = new Button("Stop");
        buttonStop.setStyle("-fx-background-color: #FF0000");
        buttonStop.setDisable(true);

        /*tFStringTimeFactor.textProperty().addListener(((observable, oldValue, newValue) -> {
            double newIntTime = 1000*Double.parseDouble(tFFlightlineLength.getText())*0.85/Double.parseDouble(newValue);
            tFIntTime.setText(Integer.toString((int) newIntTime));
        }));*/
        //Calibration
        buttonStart.setOnAction(e -> {
            //Start advancing and begin timer thread
            buttonStart.setDisable(true);
            buttonStop.setDisable(false);
            try{
                serialPort.writeString("AD+\r");
            } catch (Exception e1){
                e1.printStackTrace();
            }
            timeStart = System.currentTimeMillis();
        });
        buttonStop.setOnAction(e -> {
            //Stop advancing and write new string time
            buttonStop.setDisable(true);
            try{
                serialPort.writeString("AD\r");
            } catch (Exception e1){
                e1.printStackTrace();
            }
            timeStop = System.currentTimeMillis();
            double timeDiff = ((timeStop)-(timeStart))/1000;
            double lengthDiff = Double.parseDouble(tFFlightlineLength.getText());
            stringCalFactor = lengthDiff/timeDiff;
            tFStringTimeFactor.setText(Double.toString(lengthDiff/timeDiff));
        });
        hB5.getChildren().addAll(lB5, reg5, tFStringTimeFactor, buttonStart, buttonStop);

        rBM0.setOnMouseClicked(e->{
            lB4.setText("Sample Length (ft)");
            lB3.setText("Flightline Length (ft)");
            lB5.setText("Calibrate Speed (ft/s): ");
        });
        rBM1.setOnMouseClicked(e->{
            lB4.setText("Sample Length (m)");
            lB3.setText("Flightline Length (m)");
            lB5.setText("Calibrate Speed (m/s): ");
        });


        HBox hB6 = new HBox(10);
        Label lB6 = new Label("Direct Command: ");
        Region reg6 = new Region();
        hB6.setHgrow(reg6, Priority.ALWAYS);
        Button buttonFwd = new Button("Fwd");
        Button buttonStp = new Button("Stop");
        Button buttonRev = new Button("Rev");

        //Calibration
        buttonRev.setOnAction(e -> {
            buttonFwd.setDisable(true);
            try {
                serialPort.writeString("BD-\r");
            } catch (Exception e1) {
                e1.printStackTrace();
            }
        });
        buttonFwd.setOnAction(e -> {
            buttonRev.setDisable(true);
            try{
                serialPort.writeString("AD+\r");
            } catch (Exception e1){
                e1.printStackTrace();
            }
        });
        buttonStp.setOnAction(e -> {

            if(buttonFwd.isDisabled()) {
                try{
                    serialPort.writeString("BD\r");
                } catch (Exception e1){
                    e1.printStackTrace();
                }
            } else if (buttonRev.isDisabled()) {
                try{
                    serialPort.writeString("AD\r");
                } catch (Exception e1){
                    e1.printStackTrace();
                }
            }
            buttonRev.setDisable(false);
            buttonFwd.setDisable(false);

        });

        hB6.getChildren().addAll(lB6, reg6, buttonRev, buttonStp, buttonFwd);

        HBox hB61 = new HBox(10);
        Label lB61 = new Label("Command Line: ");
        Region reg61 = new Region();
        hB61.setHgrow(reg61, Priority.ALWAYS);
        Button buttonSend = new Button("Send");
        TextField tFCommand = new TextField();
        tFCommand.setPrefWidth(90);
        buttonSend.setOnAction(e -> {
            try{
                serialPort.writeString(tFCommand.getText()+"\r");
            } catch (Exception e1){
                e1.printStackTrace();
            }
            tFCommand.clear();
        });
        hB61.getChildren().addAll(lB61, reg61, tFCommand, buttonSend);

        //Divide it up!
        vBoxLeft.getChildren().addAll(hBoxSpectLabel, hB0 ,hBExWv, hBEmWv, hB1, hB2, hBSpectTest, hBoxStringDriveLabel, hBMetric, hB3, hB4, hB5, hB6, hB61);

        //******************************************************************************************

        //Divide it up!
        VBox vBoxRight = new VBox(10);

        HBox hBoxScannerLabel = new HBox(5);
        Text scannerLabel = new Text("WSP Processing");
        scannerLabel.setFont(Font.font(null, FontWeight.BOLD, 24));
        scannerLabel.setFill(Color.WHITESMOKE);
        hBoxScannerLabel.getChildren().addAll(scannerLabel);
        hBoxScannerLabel.setStyle("-fx-background-color: gray");

        HBox hBSpread = new HBox(10);
        Label lBSpread = new Label("Spread factor: ");
        Region regSpread = new Region();
        hBSpread.setHgrow(regSpread, Priority.ALWAYS);
        ToggleGroup tGSpread = new ToggleGroup();
        RadioButton rBSpread1 = new RadioButton("Adaptive");
        RadioButton rBSpread2 = new RadioButton("Direct");
        if(sFIdentifier==1){
            rBSpread1.setSelected(true);
            rBSpread2.setSelected(false);
        } else {
            rBSpread1.setSelected(false);
            rBSpread2.setSelected(true);
        }
        rBSpread1.setToggleGroup(tGSpread);
        rBSpread2.setToggleGroup(tGSpread);
        hBSpread.getChildren().addAll(lBSpread, regSpread, rBSpread1, rBSpread2);

        HBox hBSpreadEq = new HBox(10);
        Label lBSpreadEq = new Label();
        Region regEq1 = new Region();
        hBSpreadEq.setHgrow(regEq1, Priority.ALWAYS);
        Region regEq2 = new Region();
        hBSpreadEq.setHgrow(regEq2, Priority.ALWAYS);
        String adaptiveEq = "DD = DS/(A*DS^2+B*DS+C)";
        String directEq = "DD = A*DS^2+B*DS+C";
        if(rBSpread1.isSelected()){
            lBSpreadEq.setText(adaptiveEq);
        } else {
            lBSpreadEq.setText(directEq);
        }
        rBSpread1.setOnMouseClicked(e -> lBSpreadEq.setText(adaptiveEq));
        rBSpread2.setOnMouseClicked(e -> lBSpreadEq.setText(directEq));
        hBSpreadEq.getChildren().addAll(regEq1, lBSpreadEq, regEq2);

        HBox hBSpreadA = new HBox(10);
        Label lBSpreadA = new Label("Spread Factor A:");
        Region regSpreadA = new Region();
        hBSpreadA.setHgrow(regSpreadA, Priority.ALWAYS);
        TextField tFSpreadA = new TextField(Double.toString(spreadA));
        hBSpreadA.getChildren().addAll(lBSpreadA, regSpreadA, tFSpreadA);

        HBox hBSpreadB = new HBox(10);
        Label lBSpreadB = new Label("Spread Factor B:");
        Region regSpreadB = new Region();
        hBSpreadB.setHgrow(regSpreadB, Priority.ALWAYS);
        TextField tFSpreadB = new TextField(Double.toString(spreadB));
        hBSpreadB.getChildren().addAll(lBSpreadB, regSpreadB, tFSpreadB);

        HBox hBSpreadC = new HBox(10);
        Label lBSpreadC = new Label("Spread Factor C:");
        Region regSpreadC = new Region();
        hBSpreadC.setHgrow(regSpreadC, Priority.ALWAYS);
        TextField tFSpreadC = new TextField(Double.toString(spreadC));
        hBSpreadC.getChildren().addAll(lBSpreadC, regSpreadC, tFSpreadC);

        HBox hBThreshold = new HBox(10);
        Label lBThreshold = new Label("Default Threshold:");
        Region regThreshold = new Region();
        hBThreshold.setHgrow(regThreshold, Priority.ALWAYS);
        Slider sliderThreshold = new Slider(0,255,threshold);
        sliderThreshold.setOrientation(Orientation.HORIZONTAL);
        sliderThreshold.setMajorTickUnit(2);
        sliderThreshold.setPrefWidth(175);
        Label lBThresholdVal = new Label(Integer.toString(threshold));
        sliderThreshold.valueProperty().addListener(new ChangeListener<Number>() {
            @Override
            public void changed(ObservableValue<? extends Number> observable, Number oldValue, Number newValue) {
                lBThresholdVal.setText(Integer.toString((int) sliderThreshold.getValue()));
            }
        });
        hBThreshold.getChildren().addAll(lBThreshold, regThreshold, lBThresholdVal, sliderThreshold);

        HBox hBCrop = new HBox(10);
        Label lBCrop = new Label("Region of Interest: ");
        Region regCrop = new Region();
        hBCrop.setHgrow(regCrop, Priority.ALWAYS);
        Slider sliderCrop = new Slider(1,100,(crop*100));
        sliderCrop.setOrientation(Orientation.HORIZONTAL);
        Label lBCropVal = new Label(Integer.toString((int) (crop*100)) + "%");
        sliderCrop.valueProperty().addListener(new ChangeListener<Number>() {
            @Override
            public void changed(ObservableValue<? extends Number> observable, Number oldValue, Number newValue) {
                lBCropVal.setText(Integer.toString((int) sliderCrop.getValue()) + "%");
            }
        });
        hBCrop.getChildren().addAll(lBCrop, regCrop, lBCropVal, sliderCrop);

        HBox hBSpacing = new HBox(10);
        Label lBSpacing = new Label("Card Spacing: ");
        Region regSpacing = new Region();
        hBSpacing.setHgrow(regSpacing, Priority.ALWAYS);
        TextField tFSpacing = new TextField(Integer.toString(spacing));
        hBSpacing.getChildren().addAll(lBSpacing, regSpacing, tFSpacing);

        //ToDo AccuStain=comment
//        hBoxScannerLabel.setDisable(true);
//        hBSpread.setDisable(true);
//        hBSpreadEq.setDisable(true);
//        hBSpreadA.setDisable(true);
//        hBSpreadB.setDisable(true);
//        hBSpreadC.setDisable(true);
//        hBThreshold.setDisable(true);
//        hBCrop.setDisable(true);
//        hBSpacing.setDisable(true);


        //Label plotSettingsLabel = new Label("Other Settings");
        HBox hBoxPlotSettingsLabel = new HBox(5);
        Text plotSettingsLabel = new Text("Report");
        plotSettingsLabel.setFont(Font.font(null, FontWeight.BOLD, 24));
        plotSettingsLabel.setFill(Color.WHITESMOKE);
        hBoxPlotSettingsLabel.getChildren().addAll(plotSettingsLabel);
        hBoxPlotSettingsLabel.setStyle("-fx-background-color: gray");

        HBox hB71 = new HBox(10);
        Label lb71 = new Label("Report Logo:");
        Region reg71 = new Region();
        hB71.setHgrow(reg71, Priority.ALWAYS);
        CheckBox cBRL = new CheckBox();
        cBRL.setSelected(useLogo);
        Button logoFile = new Button(logo);
        logoFile.setMaxWidth(200);
        logoFile.setDisable(!useLogo);
        cBRL.setOnMouseClicked(e->{
            logoFile.setDisable(!cBRL.isSelected());
        });
        logoFile.setOnMouseClicked(e->{
            FileChooser fc = new FileChooser();
            fc.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("Logo Files", "*.png","*.jpg","*.tiff"));
            String newl = fc.showOpenDialog(null).toString();
            Controller.setLogo(newl);
            logoFile.setText(newl);
        });
        hB71.getChildren().addAll(lb71, reg71, logoFile, cBRL);

        HBox hB72 = new HBox(10);
        Label lB72 = new Label("Auto-Invert Pass Heading:");
        Region reg72 = new Region();
        hB72.setHgrow(reg72, Priority.ALWAYS);
        CheckBox cBAIPH = new CheckBox();
        cBAIPH.setSelected(invertPH);
        hB72.getChildren().addAll(lB72, reg72, cBAIPH);

        HBox hB73 = new HBox(10);
        Label lB73 = new Label("Show swath as: ");
        Region reg73 = new Region();
        hB73.setHgrow(reg73, Priority.ALWAYS);
        ToggleGroup tGFP = new ToggleGroup();
        RadioButton rBFPF = new RadioButton("L Wing on L");
        rBFPF.setToggleGroup(tGFP);
        RadioButton rBFPT = new RadioButton("L Wing on R");
        rBFPT.setToggleGroup(tGFP);
        if(isFlipPattern){
            tGFP.selectToggle(rBFPT);
        } else {
            tGFP.selectToggle(rBFPF);
        }
        hB73.getChildren().addAll(lB73, reg73, rBFPF, rBFPT);

        //Divide it up!
        vBoxRight.getChildren().addAll(hBoxScannerLabel, hBSpread, hBSpreadEq, hBSpreadA, hBSpreadB, hBSpreadC, hBThreshold, hBCrop, hBSpacing, hBoxPlotSettingsLabel, hB73, hB71, hB72);

        hBoxTop.getChildren().addAll(vBoxLeft, separator0, vBoxRight);

        /////////////////////////////////////////////////////////////////////////////////////////////


        HBox hB8 = new HBox(10);
        hB8.setStyle("-fx-background-color: gray");
        Region reg8 = new Region();
        hB8.setHgrow(reg8, Priority.ALWAYS);
        Region reg80 = new Region();
        hB8.setHgrow(reg80, Priority.ALWAYS);
        Region reg800 = new Region();
        hB8.setHgrow(reg800, Priority.ALWAYS);
        Region reg8000 = new Region();
        hB8.setHgrow(reg8000, Priority.ALWAYS);
        Button buttonSave = new Button("Save");
        buttonSave.setFont(Font.font(16));
        hB8.setMargin(buttonSave, new Insets(4,4,4,4));
        Button buttonCancel = new Button("Cancel");
        buttonCancel.setFont(Font.font(16));
        hB8.setMargin(buttonCancel, new Insets(4,4,4,4));
        Button buttonResetDefaults = new Button("Restore Defaults");
        buttonResetDefaults.setFont(Font.font(16));
        hB8.setMargin(buttonResetDefaults, new Insets(4,4,4,4));
        buttonResetDefaults.setOnAction(e->{
            try{
                FileInputStream fis = new FileInputStream(params);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet("Defaults");

                //Import Spectrometer Settings
                if (sh.getRow(1) != null && sh.getRow(1).getCell(1) != null) {
                    tFExWv.setText(Integer.toString((int)sh.getRow(1).getCell(1).getNumericCellValue()));
                }
                if (sh.getRow(2) != null && sh.getRow(2).getCell(1) != null) {
                    tFTargetWavelength.setText(Integer.toString((int)sh.getRow(2).getCell(1).getNumericCellValue()));
                    if(Integer.parseInt(tFTargetWavelength.getText())==575){
                        rB000.setSelected(true);
                        //rB001.setSelected(false);
                    } else if (Integer.parseInt(tFTargetWavelength.getText())==495){
                        //rB000.setSelected(false);
                        rB001.setSelected(true);
                    } else {
                        rB002.setSelected(true);
                    }
                }
                if (sh.getRow(2) != null && sh.getRow(2).getCell(1) != null) {
                    int bCW = (int) sh.getRow(2).getCell(1).getNumericCellValue();
                    rB00.setSelected(false);
                    rB01.setSelected(false);
                    rB02.setSelected(false);
                    rB03.setSelected(false);
                    if(bCW==9){
                        rB00.setSelected(true);
                    } else if(bCW==21){
                        rB01.setSelected(true);
                    } else if(bCW==41){
                        rB02.setSelected(true);
                    } else if(bCW==61){
                        rB03.setSelected(true);
                    }
                }
                if (sh.getRow(3) != null && sh.getRow(3).getCell(1) != null) {
                    tFIntTime.setText(Integer.toString((int)sh.getRow(3).getCell(1).getNumericCellValue()));
                }

                if(rBM1.isSelected()){
                    if (sh.getRow(10) != null && sh.getRow(10).getCell(2) != null) {
                        tFFlightlineLength.setText(Double.toString(sh.getRow(10).getCell(2).getNumericCellValue()));
                    }
                    if (sh.getRow(11) != null && sh.getRow(11).getCell(2) != null) {
                        tFSampleLength.setText(Double.toString(sh.getRow(11).getCell(2).getNumericCellValue()));
                    }
                    if (sh.getRow(12) != null && sh.getRow(12).getCell(2) != null) {
                        tFStringTimeFactor.setText(Double.toString(sh.getRow(12).getCell(2).getNumericCellValue()));
                    }
                } else {
                    if (sh.getRow(10) != null && sh.getRow(10).getCell(1) != null) {
                        tFFlightlineLength.setText(Integer.toString((int)sh.getRow(10).getCell(1).getNumericCellValue()));
                    }
                    if (sh.getRow(11) != null && sh.getRow(11).getCell(1) != null) {
                        tFSampleLength.setText(Double.toString(sh.getRow(11).getCell(1).getNumericCellValue()));
                    }
                    if (sh.getRow(12) != null && sh.getRow(12).getCell(1) != null) {
                        tFStringTimeFactor.setText(Double.toString(sh.getRow(12).getCell(1).getNumericCellValue()));
                    }
                }

                //Import Scanner Settings
                if (sh.getRow(1) != null && sh.getRow(1).getCell(5) != null) {
                    if(sh.getRow(1).getCell(5).getNumericCellValue() == 1){
                        rBSpread1.setSelected(true);
                        rBSpread2.setSelected(false);
                    } else {
                        rBSpread1.setSelected(false);
                        rBSpread2.setSelected(true);
                    }
                }
                if (sh.getRow(2) != null && sh.getRow(2).getCell(5) != null) {
                    tFSpreadA.setText(Double.toString(sh.getRow(2).getCell(5).getNumericCellValue()));
                }
                if (sh.getRow(3) != null && sh.getRow(3).getCell(5) != null) {
                    tFSpreadB.setText(Double.toString(sh.getRow(3).getCell(5).getNumericCellValue()));
                }
                if (sh.getRow(4) != null && sh.getRow(4).getCell(5) != null) {
                    tFSpreadC.setText(Double.toString(sh.getRow(4).getCell(5).getNumericCellValue()));
                }
                if (sh.getRow(5) != null && sh.getRow(5).getCell(5) != null) {
                    sliderThreshold.setValue((int) sh.getRow(5).getCell(5).getNumericCellValue());
                    lBThresholdVal.setText(Integer.toString((int) sh.getRow(5).getCell(5).getNumericCellValue()));
                }
                if (sh.getRow(6) != null && sh.getRow(6).getCell(5) != null) {
                    sliderCrop.setValue((int) (sh.getRow(6).getCell(5).getNumericCellValue()*100));
                    lBCropVal.setText(Integer.toString((int) (sh.getRow(6).getCell(5).getNumericCellValue()*100)) + "%");
                }
                if (sh.getRow(7) != null && sh.getRow(7).getCell(5) != null) {
                    tFSpacing.setText(Integer.toString((int) sh.getRow(7).getCell(5).getNumericCellValue()));
                }

                //Import Other Settings

                if (sh.getRow(16) != null && sh.getRow(16).getCell(1) != null) {
                    cBRL.setSelected(sh.getRow(16).getCell(1).getBooleanCellValue());
                    logoFile.setDisable(!cBRL.isSelected());
                }
                if (sh.getRow(17) != null && sh.getRow(17).getCell(1) != null) {
                    logoFile.setText(sh.getRow(17).getCell(1).getStringCellValue());
                }
                if (sh.getRow(18) != null && sh.getRow(18).getCell(1) != null) {
                    cBAIPH.setSelected(sh.getRow(18).getCell(1).getBooleanCellValue());
                }

                fis.close();
            } catch (Exception ex){
                ex.printStackTrace();
            }
        });
        buttonCancel.setOnAction(e -> {
           window.close();
        });
        buttonSave.setOnAction(e -> {
            BigDecimal remainder = BigDecimal.valueOf(Double.parseDouble(tFFlightlineLength.getText())).remainder(BigDecimal.valueOf(Double.parseDouble(tFSampleLength.getText())));
            if (remainder.doubleValue()<0.0001) {
                //Save and Exit
                //ExcitationWavelength
                Controller.setExcitationWavelength(Integer.parseInt(tFExWv.getText()));
                //Target Pixel
                Controller.setTargetWavelength(Integer.parseInt(tFTargetWavelength.getText()));
                //Boxcar Width
                if(rB00.isSelected()){Controller.setBoxcarWidth(9);}
                if(rB01.isSelected()){Controller.setBoxcarWidth(21);}
                if(rB02.isSelected()){Controller.setBoxcarWidth(41);}
                if(rB03.isSelected()){Controller.setBoxcarWidth(61);}
                //Integration Time
                Controller.setIntegrationTime(Integer.parseInt(tFIntTime.getText()));

                //Metric
                if(rBM0.isSelected()){Controller.setMetric(false);}
                if(rBM1.isSelected()){Controller.setMetric(true);}
                //Flightline Length
                Controller.setFlightlineLength(Double.parseDouble(tFFlightlineLength.getText()));
                //Sample Length
                Controller.setSampleLength(Double.parseDouble(tFSampleLength.getText()));
                //String Drive Time Factor
                Controller.setStringVelocity(Double.parseDouble(tFStringTimeFactor.getText()));
                //Scanner
                if(rBSpread1.isSelected()){Controller.setSpreadFactor(1);}
                if(rBSpread2.isSelected()){Controller.setSpreadFactor(2);}
                Controller.setSpreadABC(Double.parseDouble(tFSpreadA.getText()), Double.parseDouble(tFSpreadB.getText()), Double.parseDouble(tFSpreadC.getText()));
                Controller.setDefaultThreshold((int) sliderThreshold.getValue());
                Controller.setCropScale(sliderCrop.getValue()/100);
                Controller.setCardSpacing(Integer.parseInt(tFSpacing.getText()));
                //Logo
                Controller.setUseLogo(cBRL.isSelected());
                //Controller.setSGFilter(cBSG.isSelected());
                //Controller.setEnableExtraInfo(cBCAI.isSelected());
                Controller.setAutoInvertPH(cBAIPH.isSelected());
                if(rBFPF.isSelected()){Controller.setFlipPattern(false);}
                if(rBFPT.isSelected()){Controller.setFlipPattern(true);}
                window.close();
            } else {
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("Parameter Error");
                alert.setHeaderText("Unable to Set Chosen Values");
                alert.setContentText("Check compatibility of input types");
                alert.showAndWait();
            }
        });
        hB8.getChildren().addAll(reg8, buttonCancel, reg80, buttonResetDefaults, reg800, buttonSave, reg8000);

        VBox layout = new VBox(10);
        layout.setPadding(new Insets(4));



        //layout.getChildren().addAll(separator00, spectLabel, separator0, hB0, hB1, hB2, hB20, hB21, separator1, stringDriveLabel, separator2, hBMetric, hB3, hB4, hB5, hB6, hB61, separator3, plotSettingsLabel, separator4, hB71, hB72, hB70, hB7, hB8);
        layout.getChildren().addAll(hBoxTop, hB8);
        layout.setAlignment(Pos.CENTER);


        Scene scene = new Scene(layout);

        window.setScene(scene);
        window.showAndWait();

    }

}
