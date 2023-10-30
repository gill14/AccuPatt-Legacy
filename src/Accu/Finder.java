package Accu;

import com.itextpdf.text.Font;
import javafx.beans.property.ReadOnlyObjectProperty;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.*;
import javafx.util.Callback;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.controlsfx.control.ToggleSwitch;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTable;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Consumer;

/**
 * Created by gill14 on 11/19/18.
 */
public class Finder implements Initializable {
    @FXML
    private Accordion accordionSteps;

    @FXML
    private TitledPane titledPaneStep1;

    @FXML
    private Button buttonChooseDirectory;

    @FXML
    private Button buttonSearch;

    @FXML
    private ToggleSwitch tSHeaders;

    @FXML
    private ToggleSwitch tSPilotBusiness;

    @FXML
    private ToggleSwitch tSAircraftReg;

    @FXML
    private ToggleSwitch tSAircraftMakeModel;

    @FXML
    private ToggleSwitch tSBoomPressure;

    @FXML
    private ToggleSwitch tSTargetRate;

    @FXML
    private ToggleSwitch tSNozzles;

    @FXML
    private ToggleSwitch tSSwath;

    @FXML
    private ToggleSwitch tSPercentBoom;

    @FXML
    private ToggleSwitch tSNozzleSpacing;

    @FXML
    private ToggleSwitch tSWinglets;

    @FXML
    private ToggleSwitch tSNotes;

    @FXML
    private ChoiceBox<String> cb_filter1;

    @FXML
    private TextField tf_filter1;

    @FXML
    private ChoiceBox<String> cb_filter2;

    @FXML
    private TextField tf_filter2;

    @FXML
    private ChoiceBox<String> cb_filter3;

    @FXML
    private TextField tf_filter3;

    @FXML
    private TableView<LineEntry> table;

    @FXML
    private TableColumn<LineEntry, String> tc_h1;

    @FXML
    private TableColumn<LineEntry, String> tc_h2;

    @FXML
    private TableColumn<LineEntry, String> tc_h3;

    @FXML
    private TableColumn<LineEntry, String> tc_h4;

    @FXML
    private TableColumn<LineEntry, String> tc_pilot;

    @FXML
    private TableColumn<LineEntry, String> tc_bus;

    @FXML
    private TableColumn<LineEntry, String> tc_state;

    @FXML
    private TableColumn<LineEntry, String> tc_regnum;

    @FXML
    private TableColumn<LineEntry, String> tc_series;

    @FXML
    private TableColumn<LineEntry, String> tc_make;

    @FXML
    private TableColumn<LineEntry, String> tc_model;

    @FXML
    private TableColumn<LineEntry, String> tc_psi;

    @FXML
    private TableColumn<LineEntry, String> tc_gpa;

    @FXML
    private TableColumn<LineEntry, String> tc_noz1;

    @FXML
    private TableColumn<LineEntry, String> tc_noz1T;

    @FXML
    private TableColumn<LineEntry, String> tc_noz1Q;

    @FXML
    private TableColumn<LineEntry, String> tc_noz1O;

    @FXML
    private TableColumn<LineEntry, String> tc_noz1D;

    @FXML
    private TableColumn<LineEntry, String> tc_noz2;

    @FXML
    private TableColumn<LineEntry, String> tc_noz2T;

    @FXML
    private TableColumn<LineEntry, String> tc_noz2Q;

    @FXML
    private TableColumn<LineEntry, String> tc_noz2O;

    @FXML
    private TableColumn<LineEntry, String> tc_noz2D;

    @FXML
    private TableColumn<LineEntry, String> tc_tsw;

    @FXML
    private TableColumn<LineEntry, String> tc_psw;

    @FXML
    private TableColumn<LineEntry, String> tc_pbw;

    @FXML
    private TableColumn<LineEntry, String> tc_ns;

    @FXML
    private TableColumn<LineEntry, String> tc_winglets;

    @FXML
    private TableColumn<LineEntry, String> tc_notes;

    @FXML
    private TableColumn<LineEntry, String> tc_file;

    @FXML
    private Label labelDirectory;

    @FXML
    private Button buttonExport;

    @FXML
    private Button buttonOpen;

    @FXML
    void clickButtonChooseDirectory(ActionEvent event) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Open Resource File");
        selectedFile = directoryChooser.showDialog(null);
        if (selectedFile!=null) {
            labelDirectory.setText(selectedFile.getAbsolutePath());
            buttonSearch.setDisable(false);
        } else {
            labelDirectory.setText("No Directory Selected");
        }
    }

    @FXML
    void clickButtonSearch(ActionEvent event) {
        setResultFilters(getKey1(),tf_filter1.getText(),getKey2(),tf_filter2.getText(),getKey3(),tf_filter3.getText());
        search(selectedFile);
        setVisibleColumns(getMap());
    }

    @FXML
    void clickButtonExport(ActionEvent event) {
        if(table.getItems().size()<1){
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Error");
            alert.setHeaderText("No data to export");
            alert.setContentText("Try a new search");
            alert.showAndWait();
        } else {
            exportData();
        }
    }

    @FXML
    void clickButtonOpen(ActionEvent event) {
        if(table.getSelectionModel().getSelectedItem()==null){
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Error");
            alert.setHeaderText("No data file selected");
            alert.setContentText("Select a row entry and try again");
            alert.showAndWait();
        } else {
            this.fileToOpen = new File(table.getSelectionModel().getSelectedItem().getFile());
            //Controller.fileToOpen = this.fileToOpen;
            // get a handle to the stage
            //Stage stage = (Stage) buttonOpen.getScene().getWindow();
            // do what you have to do
            //stage.setIconified(true);
            //System.out.println("Close Request in Finder "+fileToOpen.getAbsolutePath());
            fileFromFinder.set(fileToOpen);
        }
    }

    private File directory;
    private File fileToOpen;
    private String key1;
    private String filter1;
    private String key2;
    private String filter2;
    private String key3;
    private String filter3;
    private File selectedFile;

    private HashMap<String,Boolean> boolMap;

    private final ReadOnlyObjectWrapper<File> fileFromFinder = new ReadOnlyObjectWrapper<>();

    public ReadOnlyObjectProperty<File> fileFromFinderProperty() {
        return fileFromFinder.getReadOnlyProperty();
    }

    public final File getFileFromFinder(){
        return fileFromFinderProperty().get();
    }

    @FXML
    private ObservableList<LineEntry> data = FXCollections.observableArrayList();

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        initializeChoiceBoxes();
        accordionSteps.setExpandedPane(titledPaneStep1);
        buttonSearch.setDisable(true);
        buttonChooseDirectory.requestFocus();

        tc_h1.setCellValueFactory(cellData -> cellData.getValue().header1Property());
        tc_h2.setCellValueFactory(cellData -> cellData.getValue().header2Property());
        tc_h3.setCellValueFactory(cellData -> cellData.getValue().header3Property());
        tc_h4.setCellValueFactory(cellData -> cellData.getValue().header4Property());
        tc_pilot.setCellValueFactory(cellData -> cellData.getValue().pilotProperty());
        tc_bus.setCellValueFactory(cellData -> cellData.getValue().busProperty());
        tc_state.setCellValueFactory(cellData -> cellData.getValue().stateProperty());
        tc_regnum.setCellValueFactory(cellData -> cellData.getValue().regnumProperty());
        tc_series.setCellValueFactory(cellData -> cellData.getValue().seriesProperty());
        tc_make.setCellValueFactory(cellData -> cellData.getValue().makeProperty());
        tc_model.setCellValueFactory(cellData -> cellData.getValue().modelProperty());
        tc_psi.setCellValueFactory(cellData -> cellData.getValue().psiProperty());
        tc_gpa.setCellValueFactory(cellData -> cellData.getValue().gpaProperty());
        tc_noz1T.setCellValueFactory(cellData -> cellData.getValue().noz1TProperty());
        tc_noz1Q.setCellValueFactory(cellData -> cellData.getValue().noz1QProperty());
        tc_noz1O.setCellValueFactory(cellData -> cellData.getValue().noz1OProperty());
        tc_noz1D.setCellValueFactory(cellData -> cellData.getValue().noz1DProperty());
        tc_noz2T.setCellValueFactory(cellData -> cellData.getValue().noz2TProperty());
        tc_noz2Q.setCellValueFactory(cellData -> cellData.getValue().noz2QProperty());
        tc_noz2O.setCellValueFactory(cellData -> cellData.getValue().noz2OProperty());
        tc_noz2D.setCellValueFactory(cellData -> cellData.getValue().noz2DProperty());
        tc_tsw.setCellValueFactory(cellData -> cellData.getValue().tswProperty());
        tc_psw.setCellValueFactory(cellData -> cellData.getValue().pswProperty());
        tc_pbw.setCellValueFactory(cellData -> cellData.getValue().pbwProperty());
        tc_ns.setCellValueFactory(cellData -> cellData.getValue().nsProperty());
        tc_winglets.setCellValueFactory(cellData -> cellData.getValue().wingletsProperty());
        tc_notes.setCellValueFactory(cellData -> cellData.getValue().notesProperty());
        tc_file.setCellValueFactory(cellData -> cellData.getValue().fileProperty());

        table.getSelectionModel().selectedItemProperty().addListener((obs, oldSelection, newSelection) -> {
            if (newSelection != null) {
                buttonOpen.setDisable(false);
            }
        });
    }

    private ObservableList<LineEntry> getData(){
        return data;
    }

    private void initializeChoiceBoxes(){
        for(String key: getCBMap().keySet()){
            cb_filter1.getItems().add(getCBMap().get(key));
            cb_filter2.getItems().add(getCBMap().get(key));
            cb_filter3.getItems().add(getCBMap().get(key));
        }
    }

    private HashMap<String,String> getCBMap(){
        HashMap<String,String> map = new HashMap<>();
        map.put("headers","Headers");
        map.put("pilot","Pilot");
        map.put("bus","Business");
        map.put("state","State");
        map.put("regnum","Reg. #");
        map.put("makeModel","Acft. Make/Model");
        map.put("gpa","App. Rate");
        map.put("noz","Nozzles");
        map.put("sw","Swath");
        map.put("winglets","Winglets");
        map.put("notes","Notes");
        return map;
    }

    public HashMap<String,Boolean> getMap(){
        HashMap<String,Boolean> map = new HashMap<>();
        map.put("h1",tSHeaders.isSelected());
        map.put("h2",tSHeaders.isSelected());
        map.put("h3",tSHeaders.isSelected());
        map.put("h4",tSHeaders.isSelected());
        map.put("pilot",tSPilotBusiness.isSelected());
        map.put("bus",tSPilotBusiness.isSelected());
        map.put("state",tSPilotBusiness.isSelected());
        map.put("regnum",tSAircraftReg.isSelected());
        map.put("series",tSAircraftReg.isSelected());
        map.put("make",tSAircraftMakeModel.isSelected());
        map.put("model",tSAircraftMakeModel.isSelected());
        map.put("psi",tSBoomPressure.isSelected());
        map.put("gpa",tSTargetRate.isSelected());
        map.put("noz",tSNozzles.isSelected());
        map.put("tsw",tSSwath.isSelected());
        map.put("psw",tSSwath.isSelected());
        map.put("pbw",tSPercentBoom.isSelected());
        map.put("ns",tSNozzleSpacing.isSelected());
        map.put("winglets",tSWinglets.isSelected());
        map.put("notes",tSNotes.isSelected());
        return map;
    }

    private HashMap<String,String> filterMap1;
    private HashMap<String,String> filterMap2;
    private HashMap<String,String> filterMap3;

    private void logData(File fileName){
        try {
            FileInputStream fis = new FileInputStream(fileName);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            //Get Headers
            if (wb.getNumberOfSheets()>1 && wb.getSheetAt(1).getSheetName().contains("Aircraft Data")) {
                XSSFSheet sh = wb.getSheet("Fly-In Data");
                XSSFSheet sh1 = wb.getSheet("Aircraft Data");

                HashMap<String,String> checkNullList = new HashMap<>();

                if(sh.getRow(0).getCell(0)!=null){
                    checkNullList.put("h1",sh.getRow(0).getCell(0).getStringCellValue());
                } else { checkNullList.put("h1","");}

                if(sh.getRow(1).getCell(0)!=null){
                    checkNullList.put("h2",sh.getRow(1).getCell(0).getStringCellValue());
                } else { checkNullList.put("h2","");}

                if(sh.getRow(2).getCell(0)!=null){
                    checkNullList.put("h3",sh.getRow(2).getCell(0).getStringCellValue());
                } else { checkNullList.put("h1","");}

                if(sh.getRow(3).getCell(0)!=null){
                    checkNullList.put("h4",sh.getRow(3).getCell(0).getStringCellValue());
                } else { checkNullList.put("h4","");}

                if(sh1.getRow(0).getCell(1)!=null){
                    checkNullList.put("pilot",sh1.getRow(0).getCell(1).getStringCellValue());
                } else { checkNullList.put("pilot","");}

                if(sh1.getRow(1).getCell(1)!=null){
                    checkNullList.put("bus",sh1.getRow(1).getCell(1).getStringCellValue());
                } else { checkNullList.put("bus","");}

                if(sh1.getRow(5).getCell(1)!=null){
                    checkNullList.put("state",sh1.getRow(5).getCell(1).getStringCellValue());
                } else { checkNullList.put("state","");}

                if(sh1.getRow(6).getCell(1)!=null){
                    checkNullList.put("regnum",sh1.getRow(6).getCell(1).getStringCellValue());
                } else { checkNullList.put("regnum","");}

                if(sh1.getRow(7).getCell(1)!=null){
                    checkNullList.put("series",sh1.getRow(7).getCell(1).getStringCellValue());
                } else { checkNullList.put("series","");}

                if(sh1.getRow(8).getCell(1)!=null){
                    checkNullList.put("make",sh1.getRow(8).getCell(1).getStringCellValue());
                } else { checkNullList.put("make","");}

                if(sh1.getRow(9).getCell(1)!=null){
                    checkNullList.put("model",sh1.getRow(9).getCell(1).getStringCellValue());
                } else { checkNullList.put("model","");}

                if(sh1.getRow(18).getCell(1)!=null){
                    checkNullList.put("psi",sh1.getRow(18).getCell(1).getStringCellValue());
                } else { checkNullList.put("psi","");}

                if(sh1.getRow(19).getCell(1)!=null){
                    checkNullList.put("gpa",sh1.getRow(19).getCell(1).getStringCellValue());
                } else { checkNullList.put("gpa","");}

                if(sh1.getRow(10).getCell(1)!=null){
                    checkNullList.put("noz1T",sh1.getRow(10).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz1T","");}

                if(sh1.getRow(11).getCell(1)!=null){
                    checkNullList.put("noz1Q",sh1.getRow(11).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz1Q","");}

                if(sh1.getRow(12).getCell(1)!=null){
                    checkNullList.put("noz1O",sh1.getRow(12).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz1O","");}

                if(sh1.getRow(13).getCell(1)!=null){
                    checkNullList.put("noz1D",sh1.getRow(13).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz1D","");}

                if(sh1.getRow(14).getCell(1)!=null){
                    checkNullList.put("noz2T",sh1.getRow(14).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz2T","");}

                if(sh1.getRow(15).getCell(1)!=null){
                    checkNullList.put("noz2Q",sh1.getRow(15).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz2Q","");}

                if(sh1.getRow(16).getCell(1)!=null){
                    checkNullList.put("noz2O",sh1.getRow(16).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz2O","");}

                if(sh1.getRow(17).getCell(1)!=null){
                    checkNullList.put("noz2D",sh1.getRow(17).getCell(1).getStringCellValue());
                } else { checkNullList.put("noz2D","");}

                if(sh1.getRow(20).getCell(1)!=null){
                    checkNullList.put("tsw",sh1.getRow(20).getCell(1).getStringCellValue());
                } else { checkNullList.put("tsw","");}

                if(sh1.getRow(20).getCell(2)!=null){
                    checkNullList.put("psw",sh1.getRow(20).getCell(2).getStringCellValue());
                } else { checkNullList.put("psw","");}

                if(sh1.getRow(29)!=null && sh1.getRow(29).getCell(1)!=null){
                    checkNullList.put("pbw",sh1.getRow(29).getCell(1).getStringCellValue());
                } else { checkNullList.put("pbw","");}

                if(sh1.getRow(31)!=null && sh1.getRow(31).getCell(1)!=null){
                    checkNullList.put("ns",sh1.getRow(31).getCell(1).getStringCellValue());
                } else { checkNullList.put("ns","");}

                if(sh1.getRow(32)!=null && sh1.getRow(32).getCell(1)!=null){
                    checkNullList.put("winglets",sh1.getRow(32).getCell(1).getStringCellValue());
                } else { checkNullList.put("winglets","");}

                if(sh1.getRow(33)!=null && sh1.getRow(33).getCell(1)!=null){
                    checkNullList.put("notes",sh1.getRow(33).getCell(1).getStringCellValue());
                } else { checkNullList.put("notes","");}

                boolean logit1 = false;
                boolean logit2 = false;
                boolean logit3 = false;
                if(key1.contains("aNA") || filter1.equals("")){
                    logit1 = true;
                }
                if(key2.contains("aNA") || filter2.equals("")){
                    logit2 = true;
                }
                if(key3.contains("aNA") || filter3.equals("")){
                    logit3 = true;
                }
                for(String key : filterMap1.keySet()){
                    if(!logit1 && containsIgnoreCase(checkNullList.get(key),(filterMap1.get(key)))){
                        logit1 = true;
                        break;
                    }
                }
                for(String key : filterMap2.keySet()){
                    if(!logit2 && containsIgnoreCase(checkNullList.get(key),(filterMap2.get(key)))){
                        logit2 = true;
                        break;
                    }
                }
                for(String key : filterMap3.keySet()){
                    if(!logit3 && containsIgnoreCase(checkNullList.get(key),(filterMap3.get(key)))){
                        logit3 = true;
                        break;
                    }
                }

                if (logit1 && logit2 && logit3) {
                    data.add(new LineEntry(
                            checkNullList.get("h1"), //Header 1
                            checkNullList.get("h2"), //Header 2
                            checkNullList.get("h3"), //Header 3
                            checkNullList.get("h4"), //Header 4
                            checkNullList.get("pilot"), //Pilot
                            checkNullList.get("bus"), //Business
                            checkNullList.get("state"), //State
                            checkNullList.get("regnum"), //Regnum
                            checkNullList.get("series"), //Series
                            checkNullList.get("make"), //Acft Make
                            checkNullList.get("model"), //Acft Model
                            checkNullList.get("psi"), //PSI
                            checkNullList.get("gpa"), //GPA
                            checkNullList.get("noz1T"), //Noz1T
                            checkNullList.get("noz1Q"), //Noz1Q
                            checkNullList.get("noz1O"), //Noz1O
                            checkNullList.get("noz1D"), //Noz1D
                            checkNullList.get("noz2T"), //Noz2T
                            checkNullList.get("noz2Q"), //Noz2Q
                            checkNullList.get("noz2O"), //Noz2O
                            checkNullList.get("noz2D"), //Noz2D
                            checkNullList.get("tsw"), //tsw
                            checkNullList.get("psw"), //psw
                            checkNullList.get("pbw"), //pbw
                            checkNullList.get("ns"), //ns
                            checkNullList.get("winglets"), //winglets
                            checkNullList.get("notes"), //notes
                            fileName.toString()
                    ));
                }
            }
        } catch (Exception e) {
            System.out.println("Error with: "+fileName.getName());
            e.printStackTrace();

        }

    }

    private void logDataWRK(File filename){
        HashMap<String,String> pmap = new HashMap<>();
        int indexStart = filename.getName().indexOf(".");

        if (indexStart !=-1) {
            pmap.put("series",filename.getName().substring(indexStart+1,filename.getName().length()-1));
        } else {
            pmap.put("series","");
        }
        try{
            BufferedReader in = new BufferedReader(new FileReader(filename));
            int z=0;
            String read;
            String[] line = new String[38];
            while((read = in.readLine()) != null && z < line.length) {
                if(read.length()>0) {
                    line[z] = read;
                }
                z++;
            }
            in.close();
            pmap.put("h1",line[0].replace("\"", ""));
            pmap.put("h2",line[1].replace("\"", ""));
            pmap.put("h3",line[2].replace("\"", ""));
            pmap.put("h4",line[3].replace("\"", ""));
            //4-String Length
            //5-Analysis Speed
            //6-Noz1 Flow @40 PSI
            //7-Noz2 Flow @40 PSI
            //8-# of Noz2
            pmap.put("noz1O",line[9]);
            //10-Noz2 Orif Size
            //11-Noz1 Name
            //12-Noz2 Name
            pmap.put("noz1D",line[13]);
            //14-Noz2 Def Angle
            pmap.put("bus",line[15].replace("\"", ""));
            //16-Address
            //17-Town
            pmap.put("state",line[18].replace("\"", ""));
            //19-ZIP
            //20-Phone
            pmap.put("pilot",line[21].replace("\"", ""));
            pmap.put("regnum",line[22].replace("\"", ""));
            pmap.put("model",line[23].replace("\"", ""));
            pmap.put("noz1T",line[24].replace("\"", ""));
            //25-Total # Noz
            pmap.put("psi",line[26].replace("\"", ""));
            pmap.put("gpa",line[27].replace("\"", ""));
            pmap.put("tsw",line[28].replace("\"", ""));


            pmap.put("make","");
            int numNoz2 = Integer.parseInt(line[8].replace(" ",""));
            int numNozTot = Integer.parseInt(line[25].replace("\"", "").replace(" ",""));
            int numNoz1 = numNozTot-numNoz2;
            pmap.put("noz1Q",Integer.toString(numNoz1));
            if(numNoz2==0){
                pmap.put("noz2T","");
                pmap.put("noz2D","");
                pmap.put("noz2Q","");
                pmap.put("noz2O","");
            } else {
                pmap.put("noz2T",line[24].replace("\"", ""));
                pmap.put("noz2D",line[14]);
                pmap.put("noz2Q",line[8]);
                pmap.put("noz2O",line[10]);
            }
            pmap.put("psw","");
            pmap.put("pbw","");
            pmap.put("ns","");
            pmap.put("winglets","");
            pmap.put("notes","");


            boolean logit1 = false;
            boolean logit2 = false;
            boolean logit3 = false;
            if(key1.contains("aNA") || filter1.equals("")){
                logit1 = true;
            }
            if(key2.contains("aNA") || filter2.equals("")){
                logit2 = true;
            }
            if(key3.contains("aNA") || filter3.equals("")){
                logit3 = true;
            }
            for(String key : filterMap1.keySet()){
                if(!logit1 && containsIgnoreCase(pmap.get(key),(filterMap1.get(key)))){
                    logit1 = true;
                    break;
                }
            }
            for(String key : filterMap2.keySet()){
                if(!logit2 && containsIgnoreCase(pmap.get(key),(filterMap2.get(key)))){
                    logit2 = true;
                    break;
                }
            }
            for(String key : filterMap3.keySet()){
                if(!logit3 && containsIgnoreCase(pmap.get(key),(filterMap3.get(key)))){
                    logit3 = true;
                    break;
                }
            }

            if (logit1 && logit2 && logit3) {
                data.add(new LineEntry(
                        pmap.get("h1"), //Header 1
                        pmap.get("h2"), //Header 2
                        pmap.get("h3"), //Header 3
                        pmap.get("h4"), //Header 4
                        pmap.get("pilot"), //Pilot
                        pmap.get("bus"), //Business
                        pmap.get("state"), //State
                        pmap.get("regnum"), //Regnum
                        pmap.get("series"), //Series
                        pmap.get("make"), //Acft Make
                        pmap.get("model"), //Acft Model
                        pmap.get("psi"), //PSI
                        pmap.get("gpa"), //GPA
                        pmap.get("noz1T"), //Noz1T
                        pmap.get("noz1Q"), //Noz1Q
                        pmap.get("noz1O"), //Noz1O
                        pmap.get("noz1D"), //Noz1D
                        pmap.get("noz2T"), //Noz2T
                        pmap.get("noz2Q"), //Noz2Q
                        pmap.get("noz2O"), //Noz2O
                        pmap.get("noz2D"), //Noz2D
                        pmap.get("tsw"), //tsw
                        pmap.get("psw"), //psw
                        pmap.get("pbw"), //pbw
                        pmap.get("ns"), //ns
                        pmap.get("winglets"), //winglets
                        pmap.get("notes"), //notes
                        filename.toString()
                ));
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (NumberFormatException ne) {
            System.out.println(filename);
        }
    }

    private void logDataUSDA(File filename){
        //Split the file name ([Nxxxx] [Series Letter] [Pass Number])
        String[] parts = filename.getName().split(" ");
        String reg = parts[0];
        String series = parts[1];
        String pass = parts[2];

        HashMap<String,String> pmap = new HashMap<>();

        pmap.put("regnum",reg);
        pmap.put("series",series);
        //Check for PRN File
        File prn = new File(filename.getParent()+"/Pilot Paramters.prn");
        String misc = "";
        if(prn.exists()) {
            try {
                BufferedReader in = new BufferedReader(new FileReader(prn));
                int lineNum = 0;
                String line;
                String[] entry;
                while ((line = in.readLine()) != null) {
                    //System.out.println(line);
                    entry = line.split("\t");
                    System.out.println("..."+line+"..."+entry.length);
                    if (lineNum>0 && entry[0] != null && !entry[0].isEmpty() && entry[0].contains(reg)) {
                        if (entry.length>1 && entry[1] != null) {
                            pmap.put("pilot", entry[1]);
                        }
                        if (entry.length>4 && entry[4] != null) {
                            pmap.put("state", entry[4]);
                        }
                        //Catch misc stuff
                        if (entry.length>6 && entry[6] != null) {
                            misc = misc + " " + entry[6];
                        }
                        if (entry.length>8 && entry[8] != null) {
                            misc = misc + " " + entry[8];
                        }
                        if (entry.length>9 && entry[9] != null) {
                            misc = misc + " " + entry[9];
                            System.out.println(filename+" length9"+line);
                        }
                        pmap.put("notes", misc);
                    }
                    lineNum++;
                }
                in.close();

            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("USDA Error "+filename.getName());
            }

            boolean logit1 = false;
            boolean logit2 = false;
            boolean logit3 = false;
            if(key1.contains("aNA") || filter1.equals("")){
                logit1 = true;
            }
            if(key2.contains("aNA") || filter2.equals("")){
                logit2 = true;
            }
            if(key3.contains("aNA") || filter3.equals("")){
                logit3 = true;
            }
            for(String key : filterMap1.keySet()){
                if(!logit1 && containsIgnoreCase(pmap.get(key),(filterMap1.get(key)))){
                    logit1 = true;
                    break;
                }
            }
            for(String key : filterMap2.keySet()){
                if(!logit2 && containsIgnoreCase(pmap.get(key),(filterMap2.get(key)))){
                    logit2 = true;
                    break;
                }
            }
            for(String key : filterMap3.keySet()){
                if(!logit3 && containsIgnoreCase(pmap.get(key),(filterMap3.get(key)))){
                    logit3 = true;
                    break;
                }
            }

            if (logit1 && logit2 && logit3) {
                data.add(new LineEntry(
                        "", //Header 1
                        "", //Header 2
                        "", //Header 3
                        "", //Header 4
                        pmap.get("pilot"), //Pilot
                        "", //Business
                        pmap.get("state"), //State
                        pmap.get("regnum"), //Regnum
                        pmap.get("series"), //Series
                        "", //Acft Make
                        "", //Acft Model
                        "", //PSI
                        "", //GPA
                        "", //Noz1T
                        "", //Noz1Q
                        "", //Noz1O
                        "", //Noz1D
                        "", //Noz2T
                        "", //Noz2Q
                        "", //Noz2O
                        "", //Noz2D
                        "", //tsw
                        "", //psw
                        "", //pbw
                        "", //ns
                        "", //winglets
                        pmap.get("notes"), //notes
                        filename.toString()
                ));
            }
        }
    }

    private void exportData(){
        //Create XLSX of Data
        XSSFWorkbook wb = new XSSFWorkbook();

        //Create Sheet for Report Headers
        XSSFSheet sh = wb.createSheet("Fly-In Data");
        XSSFRow row = sh.createRow(0);

        XSSFCellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        XSSFFont whiteFont = wb.createFont();
        whiteFont.setColor(new XSSFColor(java.awt.Color.WHITE));
        whiteFont.setBold(true);
        headerStyle.setFont(whiteFont);



        int k=0;
        if(boolMap.get("h1")){row.createCell(k).setCellValue("Fly-In");k++;}
        if(boolMap.get("h2")){row.createCell(k).setCellValue("Location");k++;}
        if(boolMap.get("h3")){row.createCell(k).setCellValue("Date");k++;}
        if(boolMap.get("h4")){row.createCell(k).setCellValue("Analyst");k++;}
        if(boolMap.get("pilot")){row.createCell(k).setCellValue("Pilot");k++;}
        if(boolMap.get("bus")){row.createCell(k).setCellValue("Business");k++;}
        if(boolMap.get("state")){row.createCell(k).setCellValue("State");k++;}
        if(boolMap.get("regnum")){row.createCell(k).setCellValue("Reg. #");k++;}
        if(boolMap.get("series")){row.createCell(k).setCellValue("Series");k++;System.out.println("series");}
        if(boolMap.get("make")){row.createCell(k).setCellValue("Acft. Make");k++;}
        if(boolMap.get("model")){row.createCell(k).setCellValue("Acft. Model");k++;}
        if(boolMap.get("psi")){row.createCell(k).setCellValue("PSI");k++;}
        if(boolMap.get("gpa")){row.createCell(k).setCellValue("GPA");k++;}
        if(boolMap.get("noz")){
            row.createCell(k).setCellValue("Noz1 Type");k++;
            row.createCell(k).setCellValue("Noz1 Quant.");k++;
            row.createCell(k).setCellValue("Noz1 Orif.");k++;
            row.createCell(k).setCellValue("Noz1 Def.");k++;
            row.createCell(k).setCellValue("Noz2 Type");k++;
            row.createCell(k).setCellValue("Noz2 Quant.");k++;
            row.createCell(k).setCellValue("Noz2 Orif.");k++;
            row.createCell(k).setCellValue("Noz2 Def.");k++;
        }
        if(boolMap.get("tsw")){row.createCell(k).setCellValue("Targ. SW");k++;}
        if(boolMap.get("psw")){row.createCell(k).setCellValue("Print SW");k++;}
        if(boolMap.get("pbw")){row.createCell(k).setCellValue("% BW");k++;}
        if(boolMap.get("ns")){row.createCell(k).setCellValue("Noz Spac.");k++;}
        if(boolMap.get("winglets")){row.createCell(k).setCellValue("Winglets?");k++;}
        if(boolMap.get("notes")){row.createCell(k).setCellValue("Notes");k++;}
        row.createCell(k).setCellValue("File");

        //Testing Table
        // Set which area the table should be placed in
        //AreaReference reference = new AreaReference(new CellReference(0,0),new CellReference(k,table.getItems().size()))
        //XSSFTable table = sh.createTable();
        //org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable cttable = table.getCTTable();

        for(int m=0;m<=k;m++){
            row.getCell(m).setCellStyle(headerStyle);
        }
        XSSFCellStyle regStyle = wb.createCellStyle();
        regStyle.setBorderRight(BorderStyle.THIN);
        XSSFCellStyle altStyle = wb.createCellStyle();
        altStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        altStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        altStyle.setBorderRight(BorderStyle.THIN);

        for (int i = 0; i < table.getItems().size(); i++) {
            row = sh.createRow(i + 1);
            for (int j = 0; j < table.getVisibleLeafColumns().size()-1; j++) {
                if(table.getVisibleLeafColumns().get(j).getCellData(i) != null) {
                    row.createCell(j).setCellValue(table.getVisibleLeafColumns().get(j).getCellData(i).toString());
                }
                else {
                    row.createCell(j).setCellValue("");
                }
                if((i & 1) != 0){
                    row.getCell(j).setCellStyle(altStyle);
                } else {
                    row.getCell(j).setCellStyle(regStyle);
                }
            }
        }

        for(int l=0;l<=k;l++){
            sh.autoSizeColumn(l);
        }

        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd HHmm");
        Calendar cal = Calendar.getInstance();
        System.out.println(dateFormat.format(cal.getTime()));
        try{
            File filee;
            FileChooser fc = new FileChooser();
            fc.setInitialDirectory(directory);
            fc.setInitialFileName("AccuPatt Export "+dateFormat.format(cal.getTime())+".xlsx");
            if ((filee = fc.showSaveDialog(null)) != null) {
                wb.write(new FileOutputStream(filee));
            }
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    void search(File directory){
        //Clear
        clearForNewSearch();
        //Set Directory to search within
        this.directory = directory;
        labelDirectory.setText(directory.getAbsolutePath());
        //Make a list of xlsx files within the directory
        ArrayList<File> fileList = new ArrayList<>();
        ArrayList<File> fileListWRK = new ArrayList<>();
        ArrayList<File> fileListUSDA = new ArrayList<>();
        try {
            Files.walk(Paths.get(directory.getAbsolutePath()))
                    .filter(p -> p.toString().endsWith(".xlsx"))
                    .forEach(p -> fileList.add(p.toFile()));
            Files.walk(Paths.get(directory.getAbsolutePath()))
                    .filter(p -> p.toString().endsWith("A"))
                    .forEach(p -> fileListWRK.add(p.toFile()));
            Files.walk(Paths.get(directory.getAbsolutePath()))
                    .filter(p -> p.toString().endsWith("1 .txt"))
                    .forEach(p -> fileListUSDA.add(p.toFile()));
        } catch (IOException e) {
            e.printStackTrace();
        }
        //Excluding temporary files, send each file to logData
        for(File file : fileList){
            if(!file.getName().contains("~") && !file.getName().contains("$") && !file.isDirectory()){
                logData(file);
            }
        }
        for(File file : fileListWRK){
            if(!file.getName().contains("~") && !file.isDirectory() && file.getName().startsWith("PD")){
                logDataWRK(file);
            }
        }
        for(File file : fileListUSDA){
            if(!file.getName().contains("~") && !file.isDirectory()){
                logDataUSDA(file);
            }
        }
        //After all files are put into <LineEntry> data list, add it to the table
        table.getItems().addAll(getData());
        table.getSortOrder().add(tc_series);
        table.sort();
        table.getSortOrder().clear();
        table.getSortOrder().add(tc_regnum);
        table.sort();

    }

    private void clearForNewSearch(){
        data.clear();
        table.getItems().clear();
    }

    private HashMap<String,String> populateFilterMap(String key, String filter){
        HashMap<String,String> hMap = new HashMap<>();
        if(key.contains("headers")){
            hMap.put("h1",filter);
            hMap.put("h2",filter);
            hMap.put("h3",filter);
            hMap.put("h4",filter);
        } else if (key.contains("pilot")){
            hMap.put("pilot",filter);
        } else if (key.contains("bus")){
            hMap.put("bus",filter);
        } else if (key.contains("state")){
            hMap.put("state",filter);
        } else if (key.contains("regnum")){
            hMap.put("regnum",filter);
        } else if (key.contains("makeModel")){
            hMap.put("make",filter);
            hMap.put("model",filter);
        } else if (key.contains("gpa")){
            hMap.put("gpa",filter);
        } else if (key.contains("noz")){
            hMap.put("noz1T",filter);
            hMap.put("noz1Q",filter);
            hMap.put("noz1O",filter);
            hMap.put("noz1D",filter);
            hMap.put("noz2T",filter);
            hMap.put("noz2Q",filter);
            hMap.put("noz2O",filter);
            hMap.put("noz2D",filter);
        } else if (key.contains("sw")){
            hMap.put("tsw",filter);
            hMap.put("psw",filter);
        } else if (key.contains("winglets")){
            hMap.put("winglets",filter);
        } else if (key.contains("notes")){
            hMap.put("notes",filter);
        }
        return hMap;
    }

    void setResultFilters(String key1, String filter1, String key2, String filter2, String key3, String filter3){
        filterMap1 = new HashMap<>();
        filterMap2 = new HashMap<>();
        filterMap3 = new HashMap<>();
        this.key1 = key1;
        this.filter1 = filter1;
        filterMap1 = populateFilterMap(key1,filter1);
        this.key2 = key2;
        this.filter2 = filter2;
        filterMap2 = populateFilterMap(key2,filter2);
        this.key3 = key3;
        this.filter3 = filter3;
        filterMap3 = populateFilterMap(key3,filter3);
        System.out.println("Filter Map: "+key3+ " " +filter3);
    }

    private String getKey1(){
        String key1 = "aNA";
        for (Map.Entry<String, String> entry : getCBMap().entrySet()) {
            if (entry.getValue().equals(cb_filter1.getValue())) {
                key1 = entry.getKey();
            }
        }
        System.out.println(key1);
        return key1;
    }

    private String getKey2(){
        String key2 = "aNA";
        for (Map.Entry<String, String> entry : getCBMap().entrySet()) {
            if (entry.getValue().equals(cb_filter2.getValue())) {
                key2 = entry.getKey();
            }
        }
        return key2;
    }

    private String getKey3(){
        String key3 = "aNA";
        for (Map.Entry<String, String> entry : getCBMap().entrySet()) {
            if (entry.getValue().equals(cb_filter3.getValue())) {
                key3 = entry.getKey();
            }
        }
        return key3;
    }

    void setVisibleColumns(HashMap<String,Boolean> map){
        tc_h1.setVisible(map.get("h1"));
        tc_h2.setVisible(map.get("h2"));
        tc_h3.setVisible(map.get("h3"));
        tc_h4.setVisible(map.get("h4"));
        tc_pilot.setVisible(map.get("pilot"));
        tc_bus.setVisible(map.get("bus"));
        tc_state.setVisible(map.get("state"));
        tc_regnum.setVisible(map.get("regnum"));
        tc_series.setVisible(map.get("series"));
        tc_make.setVisible(map.get("make"));
        tc_model.setVisible(map.get("model"));
        tc_psi.setVisible(map.get("psi"));
        tc_gpa.setVisible(map.get("gpa"));
        tc_noz1.setVisible(map.get("noz"));
        tc_noz2.setVisible(map.get("noz"));
        /*tc_noz1T.setVisible(map.get("noz1T"));
        tc_noz1Q.setVisible(map.get("noz1Q"));
        tc_noz1O.setVisible(map.get("noz1O"));
        tc_noz1D.setVisible(map.get("noz1D"));
        tc_noz2T.setVisible(map.get("noz2T"));
        tc_noz2Q.setVisible(map.get("noz2Q"));
        tc_noz2O.setVisible(map.get("noz2O"));
        tc_noz2D.setVisible(map.get("noz2D"));*/
        tc_tsw.setVisible(map.get("tsw"));
        tc_psw.setVisible(map.get("psw"));
        tc_pbw.setVisible(map.get("pbw"));
        tc_ns.setVisible(map.get("ns"));
        tc_winglets.setVisible(map.get("winglets"));
        tc_notes.setVisible(map.get("notes"));

        boolMap = map;
    }

    public File getFileToOpen(){
        return this.fileToOpen;
    }

    private static boolean containsIgnoreCase(String str, String searchStr)     {
        if(str == null || searchStr == null) return false;

        final int length = searchStr.length();
        if (length == 0)
            return true;

        for (int i = str.length() - length; i >= 0; i--) {
            if (str.regionMatches(true, i, searchStr, 0, length))
                return true;
        }
        return false;
    }
}
