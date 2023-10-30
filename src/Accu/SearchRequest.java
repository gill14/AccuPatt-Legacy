package Accu;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;
import java.util.ResourceBundle;

/**
 * Created by gill14 on 11/20/18.
 */
public class SearchRequest implements Initializable {

    private File selectedFile;
    private File fileToOpen;

    @FXML
    private Button buttonChooseDir;

    @FXML
    private Label labelDirectory;

    @FXML
    private Tooltip tooltipDirectory;

    @FXML
    private Button buttonSearch;

    @FXML
    private CheckMenuItem cmi_h1;

    @FXML
    private CheckMenuItem cmi_h2;

    @FXML
    private CheckMenuItem cmi_h3;

    @FXML
    private CheckMenuItem cmi_h4;

    @FXML
    private CheckMenuItem cmi_pilot;

    @FXML
    private CheckMenuItem cmi_bus;

    @FXML
    private CheckMenuItem cmi_state;

    @FXML
    private CheckMenuItem cmi_regnum;

    @FXML
    private CheckMenuItem cmi_make;

    @FXML
    private CheckMenuItem cmi_model;

    @FXML
    private CheckMenuItem cmi_psi;

    @FXML
    private CheckMenuItem cmi_gpa;

    @FXML
    private CheckMenuItem cmi_noz;

    @FXML
    private CheckMenuItem cmi_tsw;

    @FXML
    private CheckMenuItem cmi_psw;

    @FXML
    private CheckMenuItem cmi_pbw;

    @FXML
    private CheckMenuItem cmi_ns;

    @FXML
    private CheckMenuItem cmi_winglets;

    @FXML
    private CheckMenuItem cmi_notes;

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
    void clickButtonChooseDir(ActionEvent event) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("Open Resource File");
        selectedFile = directoryChooser.showDialog(null);
        labelDirectory.setText(selectedFile.getAbsolutePath());
        tooltipDirectory.setText(labelDirectory.getText());
        buttonSearch.setDisable(false);
    }

    @FXML
    void clickButtonSearch(ActionEvent event) {
        try {
            FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("Finder.fxml"));
            try {
                fxmlLoader.load();
            } catch (IOException e) {
                e.printStackTrace();
            }
            Finder controller = fxmlLoader.getController();
            Parent p = fxmlLoader.getRoot();
            Stage stage = new Stage();
            stage.setScene(new Scene(p));
            stage.setOnHidden((WindowEvent windowEvent) -> {
                // get a handle to the stage
                Stage stage1 = (Stage) buttonSearch.getScene().getWindow();
                // do what you have to do
                stage1.close();
                //System.out.println("Close Request in Search Request "+controller.getFileToOpen().getAbsolutePath());
            });
            stage.setOnShowing((WindowEvent we) -> {
                controller.setResultFilters(getKey1(),tf_filter1.getText(),getKey2(),tf_filter2.getText(),getKey3(),tf_filter3.getText());
                controller.search(selectedFile);
                controller.setVisibleColumns(getMap());
            });
            /*stage.addEventHandler(WindowEvent.WINDOW_SHOWING, new  EventHandler<WindowEvent>()
            {
                @Override
                public void handle(WindowEvent window)
                {
                    //controller.setDirectory(selectedFile);
                    controller.setResultFilters(getKey1(),tf_filter1.getText(),getKey2(),tf_filter2.getText(),getKey3(),tf_filter3.getText());
                    controller.search(selectedFile);
                    controller.setVisibleColumns(getMap());
                }
            });*/

            stage.show();
        } catch(Exception e) {
            e.printStackTrace();
        }

    }

    public void initialize(URL url, ResourceBundle rb){
        initializeChoiceBoxes();
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
        map.put("h1","Fly-In");
        map.put("h2","Location");
        map.put("h3","Date");
        map.put("h4","Analyst");
        map.put("pilot","Pilot");
        map.put("bus","Business");
        map.put("state","State");
        map.put("regnum","Reg. #");
        map.put("make","Acft. Make");
        map.put("model","Acft. Model");
        map.put("psi","Boom Press.");
        map.put("gpa","App. Rate");
        map.put("nozT","Nozzle Type");
        map.put("nozO","Orifice Size");
        map.put("nozD","Deflection");
        map.put("tsw","Target Swath");
        map.put("psw","Printed Swath");
        map.put("winglets","Winglets");
        map.put("notes","Notes");
        return map;
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
            if (entry.getValue().equals(cb_filter2.getValue())) {
                key3 = entry.getKey();
            }
        }
        return key3;
    }

    File getFileToOpen(){
        return this.fileToOpen;
    }

    public HashMap<String,Boolean> getMap(){
        HashMap<String,Boolean> map = new HashMap<>();
        map.put("h1",cmi_h1.isSelected());
        map.put("h2",cmi_h2.isSelected());
        map.put("h3",cmi_h3.isSelected());
        map.put("h4",cmi_h4.isSelected());
        map.put("pilot",cmi_pilot.isSelected());
        map.put("bus",cmi_bus.isSelected());
        map.put("state",cmi_state.isSelected());
        map.put("regnum",cmi_regnum.isSelected());
        map.put("make",cmi_make.isSelected());
        map.put("model",cmi_model.isSelected());
        map.put("psi",cmi_psi.isSelected());
        map.put("gpa",cmi_gpa.isSelected());
        map.put("noz",cmi_noz.isSelected());
        map.put("tsw",cmi_tsw.isSelected());
        map.put("psw",cmi_psw.isSelected());
        map.put("pbw",cmi_pbw.isSelected());
        map.put("ns",cmi_ns.isSelected());
        map.put("winglets",cmi_winglets.isSelected());
        map.put("notes",cmi_notes.isSelected());
        return map;
    }

}
