package Accu;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by gill14 on 6/22/2016.
 */
public class CalibrationCalc {

    //Nozzle Model File
    public static String nozzleModelFile = "NozzleModels.xlsx";

    public static ComboBox comboBoxNozzle1Type;
    public static ComboBox comboBoxNozzle1Size;
    public static ComboBox comboBoxNozzle2Type;
    public static ComboBox comboBoxNozzle2Size;

    public static void display(){
        Stage window = new Stage();
        window.initModality(Modality.APPLICATION_MODAL);
        window.setTitle("Calibration Calculator");
        window.setMinWidth(500);

        //Set up Swath/Speed/GPA HBox
        HBox hB00 = new HBox(10);
        Label swathLabel = new Label("Target Swath (FT):");
        TextField swathText = new TextField("0");
        swathText.setMaxWidth(30);
        Label speedLabel = new Label("Speed (MPH):");
        TextField speedText = new TextField("0");
        speedText.setMaxWidth(30);
        Label rateLabel = new Label("GPA:");
        TextField rateText = new TextField("0");
        rateText.setMaxWidth(30);
        Label nozQuantLabel = new Label("Nozzle Quantity:");
        TextField nozQuantText = new TextField("0");
        nozQuantText.setMaxWidth(30);
        hB00.getChildren().addAll(swathLabel,swathText,speedLabel,speedText,rateLabel,rateText,nozQuantLabel,nozQuantText);

        //Set up Boom/Nozzle Flow Rates
        HBox hB01 = new HBox(10);
        Region reg010 = new Region();
        hB01.setHgrow(reg010, Priority.ALWAYS);
        Label nozFlow = new Label("Nozzle Flow Rate (GPM):");
        Label nozFlowLabel = new Label("0.0");
        nozFlowLabel.setStyle("-fx-font-weight: bold");
        Region reg011 = new Region();
        hB01.setHgrow(reg011, Priority.ALWAYS);
        Label boomFlow = new Label("Boom Flow Rate (GPM):");
        Label boomFlowLab = new Label("0.0");
        boomFlowLab.setStyle("-fx-font-weight: bold");
        Region reg012 = new Region();
        hB01.setHgrow(reg012, Priority.ALWAYS);
        hB01.getChildren().addAll(reg010, nozFlow, nozFlowLabel, reg011, boomFlow, boomFlowLab, reg012);

        //Boom/Noz Auto Calc
        swathText.setOnAction(e -> {
            nozFlowLabel.setText(Double.toString(calcNozFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()), Double.parseDouble(nozQuantText.getText()))));
            boomFlowLab.setText(Double.toString(calcBoomFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()))));
        });
        speedText.setOnAction(e -> {
            nozFlowLabel.setText(Double.toString(calcNozFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()), Double.parseDouble(nozQuantText.getText()))));
            boomFlowLab.setText(Double.toString(calcBoomFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()))));
        });
        rateText.setOnAction(e -> {
            nozFlowLabel.setText(Double.toString(calcNozFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()), Double.parseDouble(nozQuantText.getText()))));
            boomFlowLab.setText(Double.toString(calcBoomFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()))));
        });
        nozQuantText.setOnAction(e -> {
            nozFlowLabel.setText(Double.toString(calcNozFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()), Double.parseDouble(nozQuantText.getText()))));
            boomFlowLab.setText(Double.toString(calcBoomFlow(Double.parseDouble(swathText.getText()), Double.parseDouble(speedText.getText()), Double.parseDouble(rateText.getText()))));
        });

        Separator separator0 = new Separator(Orientation.HORIZONTAL);

        //Set up Nozzle #1 HBox
        HBox hB0 = new HBox(10);
        Label noz1TypeLabel = new Label("Nozzle #1");
        comboBoxNozzle1Type = new ComboBox();
        Label noz1SizeLabel = new Label("Size:");
        comboBoxNozzle1Size = new ComboBox();
        comboBoxNozzle1Size.setMaxWidth(80);
        Label noz1QuantLabel = new Label("Quantity:");
        TextField noz1QuantText = new TextField("0");
        noz1QuantText.setMaxWidth(30);
        hB0.getChildren().addAll(noz1TypeLabel, comboBoxNozzle1Type, noz1SizeLabel, comboBoxNozzle1Size, noz1QuantLabel, noz1QuantText);

        //Set up Nozzle #2 HBox
        HBox hB1 = new HBox(10);
        Label noz2TypeLabel = new Label("Nozzle #2");
        comboBoxNozzle2Type = new ComboBox();
        Label noz2SizeLabel = new Label("Size:");
        comboBoxNozzle2Size = new ComboBox();
        comboBoxNozzle2Size.setMaxWidth(80);
        Label noz2QuantLabel = new Label("Quantity:");
        TextField noz2QuantText = new TextField("0");
        noz2QuantText.setMaxWidth(30);
        hB1.getChildren().addAll(noz2TypeLabel, comboBoxNozzle2Type, noz2SizeLabel, comboBoxNozzle2Size, noz2QuantLabel, noz2QuantText);

        //Set up Pressure-Boom Flow Rate
        HBox hB2 = new HBox(10);
        Label pressLabel = new Label("Pressure (PSI):");
        TextField pressText = new TextField();
        pressText.setText("40");
        pressText.setMaxWidth(30);
        Region reg20 = new Region();
        hB2.setHgrow(reg20, Priority.ALWAYS);
        Label boomFlowLabelLabel = new Label("Boom Flow Rate (GPM):");
        Label boomFlowLabel = new Label("0");
        boomFlowLabel.setStyle("-fx-font-weight: bold");
        Region reg21 = new Region();
        hB2.setHgrow(reg21, Priority.ALWAYS);
        hB2.getChildren().addAll(pressLabel, pressText, reg20, boomFlowLabelLabel, boomFlowLabel, reg21);

        initializeComboboxes();
        comboBoxNozzle1Type.setOnAction(e -> {
            changeNozzle1Type();
            boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText()))));
        });
        comboBoxNozzle1Size.setOnAction(e -> boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText())))));
        noz1QuantText.setOnAction(e -> boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText())))));
        comboBoxNozzle2Type.setOnAction(e -> {
            changeNozzle2Type();
            boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText()))));
        });
        comboBoxNozzle2Size.setOnAction(e -> boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText())))));
        noz2QuantText.setOnAction(e -> boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText())))));
        pressText.setOnAction(e -> boomFlowLabel.setText(Double.toString(calculateBoomFlow(Integer.parseInt(noz1QuantText.getText()), Integer.parseInt(noz2QuantText.getText()), Integer.parseInt(pressText.getText())))));

        VBox layout = new VBox(10);
        layout.getChildren().addAll(hB00, hB01, separator0, hB0, hB1, hB2);
        layout.setAlignment(Pos.CENTER);
        layout.setPadding(new Insets(4));

        Scene scene = new Scene(layout);
        window.setScene(scene);
        window.showAndWait();
    }

    public static void initializeComboboxes(){
        //Set up Nozzle #1 Type Combo Box
        List nozzle1TypeData = new ArrayList();

        try{
            FileInputStream fis = new FileInputStream(nozzleModelFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            int numOfSheets = wb.getNumberOfSheets();
            for(int i=0; i<numOfSheets; i++){
                nozzle1TypeData.add(wb.getSheetAt(i).getSheetName());
            }
            fis.close();
            ObservableList<String> nozzle1Types = FXCollections.observableArrayList(nozzle1TypeData);
            comboBoxNozzle1Type.getItems().clear();
            comboBoxNozzle1Type.setItems(nozzle1Types);
            comboBoxNozzle1Type.setValue("Select Nozzle");
        } catch (Exception e){
            e.printStackTrace();
        }

        comboBoxNozzle1Type.setEditable(true);
        comboBoxNozzle1Type.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle1Type.setValue(comboBoxNozzle1Type.getEditor().getText());
        });
        comboBoxNozzle1Size.setEditable(true);
        comboBoxNozzle1Size.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle1Size.setValue(comboBoxNozzle1Size.getEditor().getText());
        });

        //Set up Nozzle #2 Type Combo Box
        List nozzle2TypeData = new ArrayList();
        try{
            FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            int numOfSheets = wb.getNumberOfSheets();
            for(int i=0; i<numOfSheets; i++){
                nozzle2TypeData.add(wb.getSheetAt(i).getSheetName());
            }
            fis.close();
            ObservableList<String> nozzle2Types = FXCollections.observableArrayList(nozzle2TypeData);
            comboBoxNozzle2Type.getItems().clear();
            comboBoxNozzle2Type.setItems(nozzle2Types);
            comboBoxNozzle2Type.setValue("Select Nozzle");
        } catch (Exception e){
            e.printStackTrace();
        }
        comboBoxNozzle2Type.setEditable(true);
        comboBoxNozzle2Type.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle2Type.setValue(comboBoxNozzle2Type.getEditor().getText());
        });
        comboBoxNozzle2Size.setEditable(true);
        comboBoxNozzle2Size.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle2Size.setValue(comboBoxNozzle2Size.getEditor().getText());
        });

        //Set Defaults for Size/Def Boxes
        comboBoxNozzle1Size.setValue("N/A");
        comboBoxNozzle2Size.setValue("N/A");
    }

    public static void changeNozzle1Type() {
        //Set up Nozzle #1 Size Combo Box
        List nozzle1SizeData = new ArrayList();

        //New Method
        if (comboBoxNozzle1Type.getValue() != "Select Nozzle" && comboBoxNozzle1Size.getValue() != "N/A") {
            try{
                FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                XSSFRow row3 = sh.getRow(3);
                Iterator cells = row3.cellIterator();
                List cellData = new ArrayList();
                while (cells.hasNext()){
                    XSSFCell cell = (XSSFCell) cells.next();
                    if(cell.getColumnIndex() != 0){
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_STRING:
                                cellData.add(cell.getRichStringCellValue().getString());
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                cellData.add(cell.getNumericCellValue());
                                break;
                        }
                    }
                }
                nozzle1SizeData.addAll(cellData);
                fis.close();
                ObservableList<String> nozzle1Sizes = FXCollections.observableArrayList(nozzle1SizeData);
                comboBoxNozzle1Size.getItems().clear();
                comboBoxNozzle1Size.setItems(nozzle1Sizes);
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    public static void changeNozzle2Type() {
        //Set up Nozzle #2 Size Combo Box
        List nozzle2SizeData = new ArrayList();

        if (comboBoxNozzle2Type.getValue() != "Select Nozzle" && comboBoxNozzle2Size.getValue() != "N/A") {
            try{
                FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString());
                XSSFRow row3 = sh.getRow(3);
                Iterator cells = row3.cellIterator();
                List cellData = new ArrayList();
                while (cells.hasNext()){
                    XSSFCell cell = (XSSFCell) cells.next();
                    if(cell.getColumnIndex() != 0){
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_STRING:
                                cellData.add(cell.getRichStringCellValue().getString());
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                cellData.add(cell.getNumericCellValue());
                                break;
                        }
                    }
                }
                nozzle2SizeData.addAll(cellData);
                fis.close();
                ObservableList<String> nozzle2Sizes = FXCollections.observableArrayList(nozzle2SizeData);
                comboBoxNozzle2Size.getItems().clear();
                comboBoxNozzle2Size.setItems(nozzle2Sizes);
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    public static double calcBoomFlow(double swath, double speed, double GPA){
        double GPM = 0;
        if(swath > 0){
            if(speed > 0){
                if(GPA > 0){
                    GPM += (swath*speed*GPA)/495;
                }
            }
        }
        return Math.round(GPM*100000d)/100000d;
    }

    public static double calcNozFlow(double swath, double speed, double GPA, double nozQuant){
        double GPM = 0;
        if(swath > 0){
            if(speed > 0){
                if(GPA > 0){
                    if(nozQuant > 0){
                        GPM += ((swath*speed*GPA)/495)/nozQuant;
                    }
                }
            }
        }
        return Math.round(GPM*100000d)/100000d;
    }

    public static double calculateBoomFlow(int noz1Quant, int noz2Quant, int press){
        double GPM = 0;
        double orif = 0;
        int orifPos = 0;
        if (press > 0) {
            if(noz1Quant > 0){
                try{
                    FileInputStream fis = new FileInputStream(nozzleModelFile);
                    XSSFWorkbook wb = new XSSFWorkbook(fis);
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row50 = sh.getRow(50);
                    Iterator cells = row50.cellIterator();
                    //List cellData = new ArrayList();
                    while (cells.hasNext()){
                        XSSFCell cell = (XSSFCell) cells.next();
                        if(cell.getColumnIndex() != 0){
                            switch (cell.getCellType()) {
                                case XSSFCell.CELL_TYPE_STRING:
                                    orif = Double.parseDouble(cell.getRichStringCellValue().getString());
                                    break;
                                case XSSFCell.CELL_TYPE_NUMERIC:
                                    orif = cell.getNumericCellValue();
                                    break;
                            }
                            if(orif == Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString())){
                                orifPos = cell.getColumnIndex();
                            }
                        }
                    }
                    GPM += noz1Quant*(sh.getRow(51).getCell(orifPos).getNumericCellValue())*Math.sqrt(((double) press)/40.0);
                    System.out.println(press);
                    fis.close();
                } catch (Exception e){
                    e.printStackTrace();
                }
            }
            if(noz2Quant > 0){
                try{
                    FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                    XSSFWorkbook wb = new XSSFWorkbook(fis);
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row50 = sh.getRow(50);
                    Iterator cells = row50.cellIterator();
                    //List cellData = new ArrayList();
                    while (cells.hasNext()){
                        XSSFCell cell = (XSSFCell) cells.next();
                        if(cell.getColumnIndex() != 0){
                            switch (cell.getCellType()) {
                                case XSSFCell.CELL_TYPE_STRING:
                                    orif = Double.parseDouble(cell.getRichStringCellValue().getString());
                                    break;
                                case XSSFCell.CELL_TYPE_NUMERIC:
                                    orif = cell.getNumericCellValue();
                                    break;
                            }
                            if(orif == Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString())){
                                orifPos = cell.getColumnIndex();
                            }
                        }
                    }
                    GPM += noz2Quant*(sh.getRow(51).getCell(orifPos).getNumericCellValue())*Math.sqrt(((double) press)/40.0);

                    fis.close();
                } catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
        return GPM;
    }
}
