package Accu;

import Jama.LUDecomposition;
import Jama.Matrix;
import com.itextpdf.text.*;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import com.oceanoptics.omnidriver.api.wrapper.Wrapper;
import eu.gnome.morena.Device;
import eu.gnome.morena.Manager;
import eu.gnome.morena.TransferListener;
import ij.IJ;
import ij.ImagePlus;
import ij.gui.Overlay;
import ij.gui.Roi;
import ij.io.Opener;
import ij.measure.ResultsTable;
import ij.plugin.filter.EDM;
import ij.plugin.filter.ParticleAnalyzer;
import ij.plugin.frame.RoiManager;
import ij.process.AutoThresholder;
import ij.process.ImageConverter;
import ij.process.ImageProcessor;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.embed.swing.SwingFXUtils;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.*;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuItem;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.image.ImageView;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.*;
import javafx.stage.*;
import javafx.util.Duration;
import jssc.SerialPort;
import jssc.SerialPortException;
import jssc.SerialPortList;
import org.apache.commons.math3.analysis.interpolation.LinearInterpolator;
import org.apache.commons.math3.analysis.polynomials.PolynomialSplineFunction;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.DoubleStream;

public class Controller implements Initializable,TransferListener{

    //Declare MenuBar
    public Menu menuImport;
    public MenuItem menuItemOpenAny;
    public MenuItem menuItemFinder;
    public Menu menuNewSeries;
    public MenuItem menuItemNewSeries;
    public MenuItem menuItemImportAircraft;
    public MenuItem menuItemSavePDF;
    public CheckMenuItem checkMenuItemReportString;
    public CheckMenuItem checkMenuItemReportCards;
    public Menu menuSerialPort;
    public Menu menuScanners;
    public Menu menuSettings;
    public MenuItem menuItemSettings;
    public int selTogMod = 0;
    public MenuItem menuItemUserManual;
    public MenuItem menuItemCalibrationCalc;
    public MenuItem menuItemAbout;


    //Declare ScrollPane
    public ScrollPane scrollPaneMain;
    public TextField textFieldPilot;
    public TextField textFieldBusinessName;
    public TextField textFieldPhone;
    public TextField textFieldStreet;
    public TextField textFieldCity;
    public TextField textFieldState;
    public TextField textFieldZIP;
    public TextField textFieldEmail;
    public TextField textFieldRegNum;
    public TextField textFieldSeriesNum;
    public ComboBox comboBoxAircraftMake;
    public ComboBox comboBoxAircraftModel;
    public ComboBox comboBoxNozzle1Type;
    public TextField textFieldNozzle1Quant;
    public ComboBox comboBoxNozzle1Size;
    public ComboBox comboBoxNozzle1Def;
    public ComboBox comboBoxNozzle2Type;
    public TextField textFieldNozzle2Quant;
    public ComboBox comboBoxNozzle2Size;
    public ComboBox comboBoxNozzle2Def;
    public Label labelBoomPress;
    public TextField textFieldBoomPressure;
    public Label labelTargetRate;
    public TextField textFieldTargetRate;
    public Label labelTargetSwath;
    public TextField textFieldTargetSwath;
    public CheckBox cBPass1;
    public CheckBox cBPass2;
    public CheckBox cBPass3;
    public CheckBox cBPass4;
    public CheckBox cBPass5;
    public CheckBox cBPass6;
    public Label labelGroundSpeed;
    public TextField textFieldGS1;
    public TextField textFieldGS2;
    public TextField textFieldGS3;
    public TextField textFieldGS4;
    public TextField textFieldGS5;
    public TextField textFieldGS6;
    public Label labelSprayHeight;
    public TextField textFieldSH1;
    public TextField textFieldSH2;
    public TextField textFieldSH3;
    public TextField textFieldSH4;
    public TextField textFieldSH5;
    public TextField textFieldSH6;
    public TextField textFieldPH1;
    public TextField textFieldPH2;
    public TextField textFieldPH3;
    public TextField textFieldPH4;
    public TextField textFieldPH5;
    public TextField textFieldPH6;
    public TextField textFieldWD1;
    public TextField textFieldWD2;
    public TextField textFieldWD3;
    public TextField textFieldWD4;
    public TextField textFieldWD5;
    public TextField textFieldWD6;
    public Label labelWindVel;
    public TextField textFieldWV1;
    public TextField textFieldWV2;
    public TextField textFieldWV3;
    public TextField textFieldWV4;
    public TextField textFieldWV5;
    public TextField textFieldWV6;
    public Label labelTemp;
    public TextField textFieldAT;
    public TextField textFieldRH;
    //Optional Notes
    public TextField textFieldTime;
    public Label labelWingSpan;
    public TextField textFieldWingSpan;
    public Label labelBoomWidth;
    public TextField textFieldBoomWidth;
    public Label labelBoomDrop;
    public TextField textFieldBoomDrop;
    public Label labelNozzleSpacing;
    public TextField textFieldNozzleSpacing;
    public ComboBox comboBoxWinglets;
    public TextArea textAreaNotes;

    public Button buttonRun;

    //Declare TabPane
    public TabPane tabPaneMain;

    //Declare Individual Pass Tab Contents
    public NumberAxis numberAxisPassX = new NumberAxis();
    public NumberAxis numberAxisPassY = new NumberAxis();
    public AreaChart<Number,Number> areaChartPass = new AreaChart<>(numberAxisPassX, numberAxisPassY);
    public XYChart.Series seriesPass;
    public NumberAxis numberAxisPassTrimX = new NumberAxis();
    public NumberAxis numberAxisPassTrimY = new NumberAxis();
    public AreaChart<Number,Number> areaChartPassTrim = new AreaChart<>(numberAxisPassTrimX, numberAxisPassTrimY);
    public XYChart.Series seriesPassTrim;
    public Slider sliderVerticalTrim;
    public Button buttonManualReverse;
    public Label labelManualReverse;
    public Button buttonManualAdvance;
    public Label labelManualAdvance;
    public Spinner spinnerCardPassNum;
    public Spinner spinnerPassNum;
    public Spinner spinnerTrimL;
    public Spinner spinnerTrimR;
    public Button buttonPassStart;
    public Button buttonPassStop;

    //Declare Average Tab Contents
    public NumberAxis numberAxisSeriesOverlayx = new NumberAxis();
    public NumberAxis numberAxisSeriesOverlayy = new NumberAxis();
    public LineChart<Number, Number> lineChartSeriesOverlay = new LineChart<>(numberAxisSeriesOverlayx, numberAxisSeriesOverlayy);
    public NumberAxis numberAxisSeriesAveragex = new NumberAxis();
    public NumberAxis numberAxisSeriesAveragey = new NumberAxis();
    public AreaChart<Number, Number> areaChartSeriesAverage = new AreaChart<>(numberAxisSeriesAveragex, numberAxisSeriesAveragey);
    public ToggleButton toggleButtonAlignCentroid;
    public ToggleButton toggleButtonSmoothSeries;
    public ToggleButton toggleButtonPass1;
    public ToggleButton toggleButtonPass2;
    public ToggleButton toggleButtonPass3;
    public ToggleButton toggleButtonPass4;
    public ToggleButton toggleButtonPass5;
    public ToggleButton toggleButtonPass6;
    public ToggleButton toggleButtonSmoothAverage;
    public ToggleButton toggleButtonEqualizeArea;

    //Declare Swath Analysis Tab Contents
    public Label labelSwathUnitsR;
    public Label labelSwathUnitsBF;
    public NumberAxis numberAxisRacetrackx = new NumberAxis();
    public NumberAxis numberAxisRacetracky = new NumberAxis();
    public StackedAreaChart <Number, Number> areaChartRacetrack = new StackedAreaChart<>(numberAxisRacetrackx, numberAxisRacetracky);
    public GridPane gridPaneRacetrack;
    public GridPane gridPaneBackAndForth;
    public Label rt01;
    public Label rt11;
    public Label rt02;
    public Label rt12;
    public Label rt03;
    public Label rt13;
    public Label rt04;
    public Label rt14;
    public Label rt05;
    public Label rt15;
    public Label rt06;
    public Label rt16;
    public Label rt07;
    public Label rt17;
    public Label rt08;
    public Label rt18;
    public Label rt09;
    public Label rt19;
    public Label rt010;
    public Label rt110;
    public Label rt011;
    public Label rt111;
    public NumberAxis numberAxisBackAndForthx = new NumberAxis();
    public NumberAxis numberAxisBackAndForthy = new NumberAxis();
    public StackedAreaChart <Number, Number> areaChartBackAndForth = new  StackedAreaChart<>(numberAxisBackAndForthx, numberAxisBackAndForthy);
    public Label bf01;
    public Label bf11;
    public Label bf02;
    public Label bf12;
    public Label bf03;
    public Label bf13;
    public Label bf04;
    public Label bf14;
    public Label bf05;
    public Label bf15;
    public Label bf06;
    public Label bf16;
    public Label bf07;
    public Label bf17;
    public Label bf08;
    public Label bf18;
    public Label bf09;
    public Label bf19;
    public Label bf010;
    public Label bf110;
    public Label bf011;
    public Label bf111;
    public Label labelSwathFinal;
    public Label labelSwathUnits;
    public Button buttonSwathINC;
    public Button buttonSwathDEC;

    public ToggleButton toggleButtonL2;
    public ToggleButton toggleButtonL1;
    public ToggleButton toggleButtonR1;
    public ToggleButton toggleButtonR2;

    //Start AccuStain Components***************************************************************************************

    //ToDo AccuStain=true
    public Boolean AccuStain = true;
    //ToDo AccuStain=uncomment
    eu.gnome.morena.Scanner scanner;
    eu.gnome.morena.Device scanDevice;

    public Tab tabScan;
    public Tab tabDroplets;

    public Button buttonPreview;
    public Button buttonFullScan;
    public javafx.scene.control.ScrollPane scrollPaneScan;

    public VBox vBoxCard0;
    public Label labelCard0;
    public ImageView imageViewCard0;
    public Label labelVMD0;
    public Label labelDrops0;
    public Label labelCoverage0;
    public HBox hBoxCard0T;
    public ImageView imageViewCard0T;
    public Label labelCard0T;

    public VBox vBoxCard1;
    public Label labelCard1;
    public ImageView imageViewCard1;
    public Label labelVMD1;
    public Label labelDrops1;
    public Label labelCoverage1;
    public HBox hBoxCard1T;
    public ImageView imageViewCard1T;
    public Label labelCard1T;

    public VBox vBoxCard2;
    public Label labelCard2;
    public ImageView imageViewCard2;
    public Label labelVMD2;
    public Label labelDrops2;
    public Label labelCoverage2;
    public HBox hBoxCard2T;
    public ImageView imageViewCard2T;
    public Label labelCard2T;

    public VBox vBoxCard3;
    public Label labelCard3;
    public ImageView imageViewCard3;
    public Label labelVMD3;
    public Label labelDrops3;
    public Label labelCoverage3;
    public HBox hBoxCard3T;
    public ImageView imageViewCard3T;
    public Label labelCard3T;

    public VBox vBoxCard4;
    public Label labelCard4;
    public ImageView imageViewCard4;
    public Label labelVMD4;
    public Label labelDrops4;
    public Label labelCoverage4;
    public HBox hBoxCard4T;
    public ImageView imageViewCard4T;
    public Label labelCard4T;

    public VBox vBoxCard5;
    public Label labelCard5;
    public ImageView imageViewCard5;
    public Label labelVMD5;
    public Label labelDrops5;
    public Label labelCoverage5;
    public HBox hBoxCard5T;
    public ImageView imageViewCard5T;
    public Label labelCard5T;

    public VBox vBoxCard6;
    public Label labelCard6;
    public ImageView imageViewCard6;
    public Label labelVMD6;
    public Label labelDrops6;
    public Label labelCoverage6;
    public HBox hBoxCard6T;
    public ImageView imageViewCard6T;
    public Label labelCard6T;

    public VBox vBoxCard7;
    public Label labelCard7;
    public ImageView imageViewCard7;
    public Label labelVMD7;
    public Label labelDrops7;
    public Label labelCoverage7;
    public HBox hBoxCard7T;
    public ImageView imageViewCard7T;
    public Label labelCard7T;

    public VBox vBoxCard8;
    public Label labelCard8;
    public ImageView imageViewCard8;
    public Label labelVMD8;
    public Label labelDrops8;
    public Label labelCoverage8;
    public HBox hBoxCard8T;
    public ImageView imageViewCard8T;
    public Label labelCard8T;

    public Label labelCCov;
    public Label labelCVMD;
    public Label labelCDV01;
    public Label labelCDV09;
    public Label labelCRS;
    public Label labelCDSC;

    public ProgressBar progressBarScan;
    public Label labelScannerName;
    public ToggleButton togL32;
    public ToggleButton togL24;
    public ToggleButton togL16;
    public ToggleButton togL8;
    public ToggleButton tog0;
    public ToggleButton togR8;
    public ToggleButton togR16;
    public ToggleButton togR24;
    public ToggleButton togR32;


    public NumberAxis coverageYAxis = new NumberAxis();
    public NumberAxis coverageXAxis = new NumberAxis();
    public AreaChart areaChartCoverage = new AreaChart(coverageXAxis,coverageYAxis);
    public NumberAxis histYAxis = new NumberAxis();
    public CategoryAxis histXAxis = new CategoryAxis();
    public BarChart barChartDropletVolume = new BarChart(histXAxis,histYAxis);


    CardObject cardObject;
    List<Device> devices;
    String[] scannerNames;
    int scannerNum = 0;
    public static int CARD_SPACING = 8;

    static int PREVIEW_DPI = 150;
    static int FULLSCAN_DPI = 600;
    static double CROP_SCALE = 0.75;
    boolean fullScan;
    boolean ableToReRun;
    boolean validateScannersList;
    static int DEFAULT_THRESHOLD;
    public static int useSpreadFactor = 1; //1=Adaptive; 2=Direct
    public static double spreadA;
    public static double spreadB;
    public static double spreadC;


    //End AccuStain Components*****************************************************************************************


    //Declare Status Label
    public Label labelCurrentFile;
    public Label labelCom;
    public Label labelSpectrometer;

    //Metric Compatible Labeling of Units:
    public static String uAirSpeed = "MPH";
    public static String uWindSpeed = "MPH";
    public static String uPressure = "PSI";
    public static String uTemp = "Deg F";
    public static String uAppRate = "GPA";
    public static String uFlowRate = "GPM";
    public static String uHeight = "FT";
    public static String uHeightSmall = "IN";

    //Declare Current Directory
    public String version = "1.04+ Demo";
    private File currentDirectory;
    private File reportFile;
    public File currentDataFile;
    public static File fileToOpen;

    //Declare Stepper Driver
    public SerialPort serialPort;
    public String[] portNames;
    public Boolean isReversing = false;
    public Boolean isAdvancing = false;

    //Declare Spectrometer
    public static Wrapper wrapper;
    public static int SCANS_TO_AVERAGE = 1;
    public static int TARGET_WAVELENGTH = 575;
    public static int TARGET_PIXEL = 642;
    public static int EXCITATION_WAVELENGTH = 535;
    public static int EXCITATION_PIXEL = 410;
    public static int BOXCAR_WIDTH = 41;
    public static int INTEGRATION_TIME = 100000;
    public static int IS_CORRECT_FOR_ELECTRICAL_DARK = 0; //1-True, 0-False

    //Declare Timelines
    public Timeline animation1;
    public Timeline animation2;
    public Timeline animation3;
    public Timeline animation4;
    public Timeline animation5;
    public Timeline animation6;
    public double location;

    //Metric Compatability
    public static boolean METRIC = false;

    //String Drive Constants
    public static double FLIGHTLINE_LENGTH;
    public static double SAMPLE_LENGTH;
    public static double STRING_TIME;
    public static double STRING_VEL;

    //WRK Constants

    //File Names
    public static String paramsWorkbook = "Parameters.xlsx";
    public static String aircraftWorkbook = "AgAircraftData.xlsx";
    public static String nozzleWorkbook = "NozzleModels.xlsx";
    public static String userManual = "UserManual.pdf";
    public static Boolean useLogo = false;
    public static String logoFile = "SAFELogo.jpg";

    //Test Stuff
    public PatternObject patternObject;
    public int cycleNum;
    public static Boolean isFlipPattern = false;

    public ObservableList<Integer> activePasses = FXCollections.observableArrayList();
    public boolean isScannedPass1 = false;
    public boolean isScannedPass2 = false;
    public boolean isScannedPass3 = false;
    public boolean isScannedPass4 = false;
    public boolean isScannedPass5 = false;
    public boolean isScannedPass6 = false;
    public boolean isAlignCentroid = true;
    public boolean isSmoothSeries = true;
    public boolean isSmoothAverage = true;
    public boolean isSwathBox = true;
    public boolean isEqualizeArea = true;
    public boolean pass1Mark = false;
    public boolean pass2Mark = false;
    public boolean pass3Mark = false;
    public boolean pass4Mark = false;
    public boolean pass5Mark = false;
    public boolean pass6Mark = false;
    public Document pdfDocument;
    public static String header1;
    public static String header2;
    public static String header3;
    public static String header4;
    public static Boolean invertPH = false;
    public Boolean cV3 = true;
    public Boolean cV5 = true;

    @Override
    public void initialize(URL location, ResourceBundle resources) {


        //ToDo AccuStain=comment
            /*tabPaneMain.getTabs().get(4).setDisable(true);
            tabPaneMain.getTabs().get(3).setDisable(true);
            checkMenuItemReportCards.setDisable(true);
            checkMenuItemReportCards.setSelected(false);
            menuScanners.setDisable(true);*/


        importParamsWorkbook();

        //Populate Serial Port List
        refreshSerialPortList();

        //ToDo AccuStain=uncomment
        //Populate Scanners List
        Manager.getInstance();
        refreshScannersList();
        validateScannersList = false;
        fullScan = false;
        setCardSpacing();
        //

        //MenuBar Defaults
        buttonRun.setDisable(true);
        //menuItemImportAircraft.setDisable(true);

        //Set default status for Drive Booleans
        isReversing = false;
        isAdvancing = false;

        //Initialize ScrollPane ComboBoxes
        initializeComboBoxes();

        //Units
        updateLabelUnits();

        //Sync Pass Heading
        textFieldPH1.textProperty().addListener((observable, oldValue, newValue) -> {
            if (cBPass2.isSelected()) {
                textFieldPH2.setText(newValue);
            }
            if (cBPass3.isSelected()) {
                textFieldPH3.setText(newValue);
            }
            if (cBPass4.isSelected()) {
                textFieldPH4.setText(newValue);
            }
            if (cBPass5.isSelected()) {
                textFieldPH5.setText(newValue);
            }
            if (cBPass6.isSelected()) {
                textFieldPH6.setText(newValue);
            }
        });

        //Initialize AreaCharts
        areaChartPass.setLegendVisible(false);
        areaChartPass.setCreateSymbols(false);
        areaChartPass.setAnimated(false);
        areaChartPassTrim.setLegendVisible(false);
        areaChartPassTrim.setCreateSymbols(false);
        areaChartPassTrim.setAnimated(false);
        areaChartPassTrim.setId("trimmedPass-Plot");
        lineChartSeriesOverlay.setLegendVisible(false);
        lineChartSeriesOverlay.setCreateSymbols(false);
        areaChartSeriesAverage.setCreateSymbols(false);
        areaChartSeriesAverage.setId("average-Plot");
        areaChartRacetrack.setLegendVisible(false);
        areaChartRacetrack.setCreateSymbols(false);
        areaChartRacetrack.setId("cv-Plot");
        numberAxisRacetrackx.setAutoRanging(false);
        areaChartBackAndForth.setLegendVisible(false);
        areaChartBackAndForth.setCreateSymbols(false);
        areaChartBackAndForth.setId("cv-Plot");
        numberAxisBackAndForthx.setAutoRanging(false);
        resetHorizontalAxes(FLIGHTLINE_LENGTH);

        //Set Panes Disabled By Default
        scrollPaneMain.setDisable(true);
        tabPaneMain.setDisable(true);

        //Setup Card Pass Spinner
        SpinnerValueFactory<Integer> valueFactoryCPN = new SpinnerValueFactory.IntegerSpinnerValueFactory(1,6,3);
        spinnerCardPassNum.setValueFactory(valueFactoryCPN);

        //Setup Pass Spinner
        SpinnerValueFactory<Integer> valueFactoryPN = new SpinnerValueFactory.ListSpinnerValueFactory<>(activePasses);
        spinnerPassNum.setValueFactory(valueFactoryPN);

        //Setup L and R Trim Spinners
        SpinnerValueFactory.IntegerSpinnerValueFactory valueFactoryTPL = new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 200, 0);
        valueFactoryTPL.setAmountToStepBy(2);
        spinnerTrimL.setValueFactory(valueFactoryTPL);
        SpinnerValueFactory.IntegerSpinnerValueFactory valueFactoryTPR = new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 200, 0);
        valueFactoryTPR.setAmountToStepBy(2);
        spinnerTrimR.setValueFactory(valueFactoryTPR);

        //Set Default For Pass Visibility Controls
        toggleButtonPass1.setDisable(true);
        toggleButtonPass2.setDisable(true);
        toggleButtonPass3.setDisable(true);
        toggleButtonPass4.setDisable(true);
        toggleButtonPass5.setDisable(true);
        toggleButtonPass6.setDisable(true);

        //Set Default Directory
        //currentDirectory = System.getProperty("user.home");

        //Set Headers
        header1 = "(Year),(Fly-In Name)";
        header2 = "(Location)";
        header3 = "(Date)";
        header4 = "(Analyst)";

        buttonRun.getStyleClass().add("runButton");

        //Set Style for Card Buttons
        togL32.getStyleClass().add("cardToggleButton");
        togL24.getStyleClass().add("cardToggleButton");
        togL16.getStyleClass().add("cardToggleButton");
        togL8.getStyleClass().add("cardToggleButton");
        tog0.getStyleClass().add("cardToggleButton");
        togR8.getStyleClass().add("cardToggleButton");
        togR16.getStyleClass().add("cardToggleButton");
        togR24.getStyleClass().add("cardToggleButton");
        togR32.getStyleClass().add("cardToggleButton");

    }

    //File Menu

    public void clickMenuItemOpenAny(){
        File datFile;
        try{
            FileChooser fc = new FileChooser();
            if(currentDirectory != null) {
                fc.setInitialDirectory(currentDirectory);
            }
            ArrayList<String> exts = new ArrayList<>();
            exts.add(0,"*.xlsx");
            exts.add(1,"*1 .txt");
            for(int i=2; i<50; i++){
                exts.add(i,"*."+(i-1)+"A");
            }
            FileChooser.ExtensionFilter stringFiles = new FileChooser.ExtensionFilter("String Files",exts);
            fc.getExtensionFilters().add(stringFiles);
            if ((datFile = fc.showOpenDialog(null)) != null) {
                if (datFile.getAbsolutePath().endsWith(".xlsx")){
                    selTogMod = 0;
                    openAccuPatt(datFile);
                } else if (datFile.getAbsolutePath().endsWith("A")){
                    selTogMod = 1;
                    buttonRun.setDisable(true);
                    openWRK(datFile);
                } else if (datFile.getAbsolutePath().endsWith(".txt")){
                    selTogMod = 2;
                    buttonRun.setDisable(true);
                    openUSDA(datFile);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("File Open Error");
            alert.setHeaderText("Error Opening File");
            alert.setContentText("Check file contents and try again.");
            alert.showAndWait();
        }
    }

    public void clickMenuItemFinder(){
        try {
            FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("Finder.fxml"));
            Parent root1 = (Parent) fxmlLoader.load();
            Finder finder = fxmlLoader.getController();
            finder.fileFromFinderProperty().addListener((obs, oldSelection, newSelection) -> {
                System.out.println("New Selection"+newSelection.getAbsolutePath());
                if (newSelection.getAbsolutePath().endsWith(".xlsx")){
                    selTogMod = 0;
                    openAccuPatt(newSelection);
                } else if (newSelection.getAbsolutePath().endsWith("A")){
                    selTogMod = 1;
                    buttonRun.setDisable(true);
                    openWRK(newSelection);
                } else if (newSelection.getAbsolutePath().endsWith(".txt")){
                    selTogMod = 2;
                    buttonRun.setDisable(true);
                    openUSDA(newSelection);
                } else {
                    System.out.println("Invalid File Type");
                }
            });

            Stage stage = new Stage();
            stage.setScene(new Scene(root1));
            stage.show();
        } catch(Exception e) {
            e.printStackTrace();
        }
    }

    //Analysis

    public void clickButtonRun(){
        if (buttonRun.getText().equals("Save and Proceed")) {

            //Open Channel to Spectrometer
            try {
                //ToDo - Comment for MAC testing
                setupSpectrometer();
            } catch (Exception e) {
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("Peripheral Error");
                alert.setHeaderText("Unable to locate Spectrometer");
                alert.setContentText("Check Device Manager for winUSB and try again.");
                alert.showAndWait();
                return;
            }
            //Set swath to target swath (or default value)
            if(!textFieldTargetSwath.getText().isEmpty()) {
                if (Double.parseDouble(textFieldTargetSwath.getText())<FLIGHTLINE_LENGTH) {
                    labelSwathFinal.setText(textFieldTargetSwath.getText());
                } else {
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Swath Width Error");
                    alert.setHeaderText("Target Swath is greater than Flightline Length");
                    alert.setContentText("Reduce Target Swath or increase Flightline Length and Try Again");
                    alert.showAndWait();
                    return;
                }
            } else {
                if (!METRIC) {
                    labelSwathFinal.setText("65");
                } else {
                    labelSwathFinal.setText("20");
                }
            }

            if(serialPort == null){
                try {
                    if(portNames.length>0) {
                        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                        alert.setTitle("Confirmation");
                        alert.setHeaderText("Auto-selected the following COM port:");
                        alert.setContentText(portNames[portNames.length - 1]);
                        Optional<ButtonType> result = alert.showAndWait();
                        if (result.isPresent() && result.get() == ButtonType.OK) {
                            initializeSerialPort(portNames[portNames.length - 1], false);
                            CheckMenuItem cMI = (CheckMenuItem) menuSerialPort.getItems().get(portNames.length-1);
                            cMI.setSelected(true);
                        }
                    } else {
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setTitle("Peripheral Error");
                        alert.setHeaderText("No Serial Port Found");
                        alert.setContentText("Check connections, then go to Settings -> Serial Port and refresh the list.");
                        alert.showAndWait();
                        return;
                    }
                } catch (Exception e){
                    System.out.println(e);
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Peripheral Error");
                    alert.setHeaderText("No Serial Port Found");
                    alert.setContentText("Check connections, then go to Settings -> Serial Port and refresh the list.");
                    alert.showAndWait();
                    return;
                }
            }
            //If serial port is opened, let's get this show on the road
            if (serialPort != null && serialPort.isOpened()){
                buttonRun.setDisable(true);
                scrollPaneMain.setDisable(true);
                tabPaneMain.setDisable(false);
                saveCurrentAircraftAndSeriesData();
                pass1Mark = false;
                pass2Mark = false;
                pass3Mark = false;
                pass4Mark = false;
                pass5Mark = false;
                pass6Mark = false;
                updateActivePasses();
                if(cBPass6.isSelected()){
                    spinnerPassNum.getValueFactory().setValue(6);
                } else if(cBPass5.isSelected()){
                    spinnerPassNum.getValueFactory().setValue(5);
                } else if(cBPass4.isSelected()){
                    spinnerPassNum.getValueFactory().setValue(4);
                } else if(cBPass3.isSelected()){
                    spinnerPassNum.getValueFactory().setValue(3);
                } else if(cBPass2.isSelected()){
                    spinnerPassNum.getValueFactory().setValue(2);
                } else {
                    spinnerPassNum.getValueFactory().setValue(1);
                }
                updateVisibilityAnalyzing(false, (int) spinnerPassNum.getValue());

            }
        } else {
            updateDataFile();
        }
    }

    private void setupSpectrometer(){
        wrapper = new Wrapper();
        wrapper.closeAllSpectrometers();
        try{
            int num = wrapper.openAllSpectrometers();
            if (num>0) {
                wrapper.setBoxcarWidth(0,(BOXCAR_WIDTH-1)/2);
                wrapper.setScansToAverage(0, SCANS_TO_AVERAGE);
                wrapper.setIntegrationTime(0,INTEGRATION_TIME);
                wrapper.setCorrectForElectricalDark(0, IS_CORRECT_FOR_ELECTRICAL_DARK);
                double [] wav = wrapper.getWavelengths(0);
                for(int i=0; i<wav.length;i++){
                    if (Math.round(wav[i])==TARGET_WAVELENGTH){
                        TARGET_PIXEL=i;
                    }
                    if (Math.round(wav[i])==EXCITATION_WAVELENGTH){
                        EXCITATION_PIXEL=i;
                    }
                }
                labelSpectrometer.setText(wrapper.getName(0)+" EM:"+Integer.toString((int) Math.round(wav[TARGET_PIXEL]))+"("+Integer.toString(TARGET_PIXEL)+") EX:"+Integer.toString((int) Math.round(wav[EXCITATION_PIXEL]))+"("+Integer.toString(EXCITATION_PIXEL)+")");
            } else {
                labelSpectrometer.setText("None Found");
            }
        } catch (Exception e) {
            labelSpectrometer.setText("None Found");
        }
    }

    public void clickMenuItemNewSeries(){
        selTogMod = 0;
        buttonRun.setDisable(false);
        buttonRun.setText("Save and Proceed");
        scrollPaneMain.setDisable(false);
        tabPaneMain.getSelectionModel().select(0);
        tabPaneMain.setDisable(true);
        clearAircraftAndSeriesData();
        clearPatternData();
        //Default enable 3 passes
        cBPass1.setSelected(true);
        cBPass2.setSelected(true);
        cBPass3.setSelected(true);
        clearCardSlate();
        patternObject = new PatternObject();
        cardObject = new CardObject(DEFAULT_THRESHOLD);
        cardObject.getRm().setVisible(false);
    }

    public void clickMenuItemImportAircraft(){
        clickMenuItemNewSeries();
        FileChooser fc = new FileChooser();
        if (getCurrentDirectory()!=null) {
            fc.setInitialDirectory(getCurrentDirectory());
        }
        fc.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("AccuPatt Files", "*.xlsx"));
        File selectedFile = fc.showOpenDialog(null);
        currentDirectory = selectedFile.getParentFile();
        try{
            FileInputStream fis = new FileInputStream(selectedFile);
            XSSFWorkbook  wb = new XSSFWorkbook(fis);
            //Get Aircraft Data
            XSSFSheet sheet = wb.getSheet("Aircraft Data");
            //Determine if metric or not
            if (sheet.getRow(35).getCell(1) != null) {
                setMetric(sheet.getRow(35).getCell(1).getBooleanCellValue());
                updateLabelUnits();
            } else {
                METRIC = false;
            }
            if (sheet.getRow(0).getCell(1) != null) {
                textFieldPilot.setText(sheet.getRow(0).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(0).getCell(2) != null) {
                textFieldEmail.setText(sheet.getRow(0).getCell(2).getStringCellValue());
            }
            if (sheet.getRow(1).getCell(1) != null) {
                textFieldBusinessName.setText(sheet.getRow(1).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(2).getCell(1) != null) {
                textFieldPhone.setText(sheet.getRow(2).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(3).getCell(1) != null) {
                textFieldStreet.setText(sheet.getRow(3).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(4).getCell(1) != null) {
                textFieldCity.setText(sheet.getRow(4).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(5).getCell(1) != null) {
                textFieldState.setText(sheet.getRow(5).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(5).getCell(2) != null) {
                textFieldZIP.setText(sheet.getRow(5).getCell(2).getStringCellValue());
            }
            if (sheet.getRow(6).getCell(1) != null) {
                textFieldRegNum.setText(sheet.getRow(6).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(7).getCell(1) != null) {
                textFieldSeriesNum.setText(Integer.toString(Integer.parseInt(sheet.getRow(7).getCell(1).getStringCellValue())+1));
            }
            if (sheet.getRow(8).getCell(1) != null) {
                comboBoxAircraftMake.setValue(sheet.getRow(8).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(9).getCell(1) != null) {
                comboBoxAircraftModel.setValue(sheet.getRow(9).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(10).getCell(1) != null) {
                comboBoxNozzle1Type.setValue(sheet.getRow(10).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(11).getCell(1) != null) {
                textFieldNozzle1Quant.setText(sheet.getRow(11).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(12).getCell(1) != null) {
                comboBoxNozzle1Size.setValue(sheet.getRow(12).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(13).getCell(1) != null) {
                comboBoxNozzle1Def.setValue(sheet.getRow(13).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(14).getCell(1) != null) {
                comboBoxNozzle2Type.setValue(sheet.getRow(14).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(15).getCell(1) != null) {
                textFieldNozzle2Quant.setText(sheet.getRow(15).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(16).getCell(1) != null) {
                comboBoxNozzle2Size.setValue(sheet.getRow(16).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(17).getCell(1) != null) {
                comboBoxNozzle2Def.setValue(sheet.getRow(17).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(18).getCell(1) != null) {
                textFieldBoomPressure.setText(sheet.getRow(18).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(19).getCell(1) != null) {
                textFieldTargetRate.setText(sheet.getRow(19).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(20).getCell(1) != null) {
                textFieldTargetSwath.setText(sheet.getRow(20).getCell(1).getStringCellValue());
                labelSwathFinal.setText(sheet.getRow(20).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(26).getCell(1) != null) {
                textFieldTime.setText(sheet.getRow(26).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(27) != null && sheet.getRow(27).getCell(1) != null) {
                textFieldWingSpan.setText(sheet.getRow(27).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(28) != null && sheet.getRow(28).getCell(1) != null) {
                textFieldBoomWidth.setText(sheet.getRow(28).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(30) != null && sheet.getRow(30).getCell(1) != null) {
                textFieldBoomDrop.setText(sheet.getRow(30).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(31) != null && sheet.getRow(31).getCell(1) != null) {
                textFieldNozzleSpacing.setText(sheet.getRow(31).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(32) != null && sheet.getRow(32).getCell(1) != null) {
                comboBoxWinglets.setValue(sheet.getRow(32).getCell(1).getStringCellValue());
            }
            if (sheet.getRow(33) != null && sheet.getRow(33).getCell(1) != null) {
                textAreaNotes.setText(sheet.getRow(33).getCell(1).getStringCellValue());
            }
            fis.close();
            } catch(Exception e){e.printStackTrace();}
    }                       //XLSX

    //Report Menu

    public void editHeaders(){
        HeaderEditor.display(header1, header2, header3, header4);
        updateParamsWorkbook();
        updateHeadersOnDataFile();
        clickMenuItemGenerateReport();
    }

    public void updateHeadersOnDataFile(){
        //Save Pattern Data to Current Data File
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("Fly-In Data");
            if (header1 != null && !header1.isEmpty()) {
                sh.createRow(0).createCell(0).setCellValue(header1);
            } else {
                sh.createRow(0).createCell(0);
            }
            if (header2 != null && !header2.isEmpty()) {
                sh.createRow(1).createCell(0).setCellValue(header2);
            } else {
                sh.createRow(1).createCell(0);
            }
            if (header3 != null && !header3.isEmpty()) {
                sh.createRow(2).createCell(0).setCellValue(header3);
            } else {
                sh.createRow(2).createCell(0);
            }
            if (header4 != null && !header4.isEmpty()) {
                sh.createRow(3).createCell(0).setCellValue(header4);
            } else {
                sh.createRow(3).createCell(0);
            }

            sh.autoSizeColumn(0);

            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}

        //Update SAFE Report as well
        //File reportFile = new File(currentDirectory+"/SAFEReport.xlsx");
        if (reportFile!=null && reportFile.exists()) {
            try {
                FileInputStream fis = new FileInputStream(reportFile);
                Workbook wb = WorkbookFactory.create(fis);
                Sheet sh = wb.getSheet("Fly-In Data");

                sh.getRow(0).getCell(2).setCellValue(header1);
                sh.getRow(1).getCell(2).setCellValue(header2);
                sh.getRow(2).getCell(2).setCellValue(header3);
                sh.getRow(3).getCell(2).setCellValue(header4);

                FileOutputStream fos = new FileOutputStream(reportFile);
                wb.write(fos);
                fos.close();
            } catch (Exception e) {
                e.printStackTrace();
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("File Save Error");
                alert.setHeaderText("Unable to Write File");
                alert.setContentText("Check if file is open in another program and try again.");
                alert.showAndWait();
            }
        }

    }                           //XLSX

    public void clickMenuItemGenerateReport(){


        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("Generate Report");
        alert.setHeaderText("Confirm Report Headers");
        alert.setContentText(header1+"\n"+header2+"\n"+header3+"\n"+header4);

        ButtonType buttonGenerate = new ButtonType("Generate");
        ButtonType buttonEditHeaders = new ButtonType("Edit Headers");
        ButtonType buttonTypeCancel = new ButtonType("Cancel", ButtonBar.ButtonData.CANCEL_CLOSE);

        alert.getButtonTypes().setAll(buttonGenerate, buttonEditHeaders, buttonTypeCancel);

        Optional<ButtonType> result = alert.showAndWait();
        if (result.get() == buttonGenerate){
            savePrintedSwathWidth();
            String name;
            if(textFieldSeriesNum.getLength() == 1) {
                name = textFieldRegNum.getText() + " 0" + textFieldSeriesNum.getText() + ".pdf";
            } else if(textFieldSeriesNum.getLength() > 1){
                name = textFieldRegNum.getText() + " " + textFieldSeriesNum.getText() + ".pdf";
            } else {
                name = textFieldRegNum.getText() + ".pdf";
            }
            generatePDF(new File (currentDirectory+File.separator+name), checkMenuItemReportString.isSelected(), checkMenuItemReportCards.isSelected());
        } else if (result.get() == buttonEditHeaders) {
            editHeaders();
        }
    }

    //IO Menu

    public void initializeSerialPort(String port, Boolean isOpened) {
        try{
            if(serialPort != null && serialPort.isOpened()){
                serialPort.closePort();
            }
            //serialPort = new SerialPort(port);
            if(isOpened){
                if(serialPort.isOpened()){
                    serialPort.closePort();
                }
                labelCom.setText("");
                //hBoxDrivePass1.setDisable(true);
                //hBoxDrivePass2.setDisable(true);
                //hBoxDrivePass3.setDisable(true);
            } else if(!isOpened){
                serialPort = new SerialPort(port);

                serialPort.openPort();
                serialPort.setParams(9600, 8, 1, 0);
                labelCom.setText(port);
                //hBoxDrivePass1.setDisable(false);
                //hBoxDrivePass2.setDisable(false);
                //hBoxDrivePass3.setDisable(false);
            }

        } catch (SerialPortException e) {
            e.printStackTrace();
        }
    }

    public void clickMenuItemSerialPortListRefresh(){
        labelCom.setText("");
        refreshSerialPortList();
    }

    public void refreshSerialPortList(){
        menuSerialPort.getItems().clear();
        //portNames = Arrays.asArrayList(SerialPortList.getPortNames());
        portNames = new String[SerialPortList.getPortNames().length];
        portNames = SerialPortList.getPortNames();

            for(int i=0; i<portNames.length; i++){
                switch(i) {
                    case (0):
                        CheckMenuItem checkMenuItemPort0 = (new CheckMenuItem(portNames[0]));
                        menuSerialPort.getItems().add(checkMenuItemPort0);
                        checkMenuItemPort0.setOnAction((event -> {
                            if(checkMenuItemPort0.isSelected()){
                                initializeSerialPort(portNames[0], false);
                            } else if(!checkMenuItemPort0.isSelected()){
                                initializeSerialPort(portNames[0], true);
                            }
                        }));
                        break;
                    case (1):
                        CheckMenuItem checkMenuItemPort1 = new CheckMenuItem(portNames[1]);
                        menuSerialPort.getItems().add(checkMenuItemPort1);
                        checkMenuItemPort1.setOnAction((event -> {
                            if(checkMenuItemPort1.isSelected()){
                                initializeSerialPort(portNames[1], false);
                            } else if(!checkMenuItemPort1.isSelected()){
                                initializeSerialPort(portNames[1], true);
                            }
                        }));
                        break;

                    case (2):
                        CheckMenuItem checkMenuItemPort2 = new CheckMenuItem(portNames[2]);
                        menuSerialPort.getItems().add(checkMenuItemPort2);
                        checkMenuItemPort2.setOnAction((event -> {
                            if(checkMenuItemPort2.isSelected()){
                                initializeSerialPort(portNames[2], false);
                            } else if(!checkMenuItemPort2.isSelected()){
                                initializeSerialPort(portNames[2], true);
                            }
                        }));
                        break;

                }
            }
    }

    public void initialValidateScannersList(){
        if (AccuStain) {
            //ToDo AccuStain=uncomment
            if(!validateScannersList){
                refreshScannersList();
                validateScannersList = true;
            }
        }
    }

    public void refreshScannersList(){
        menuScanners.getItems().clear();
        Manager manager = Manager.getInstance();
        System.out.println(manager.toString());
        devices = manager.listDevices();
        scannerNames = new String[devices.size()];
        for(int i=0; i<devices.size(); i++){
            scannerNames[i] = devices.get(i).toString();
            System.out.println(scannerNames[i]);
            switch(i) {
                case (0):
                    CheckMenuItem checkMenuItemScanner0 = (new CheckMenuItem(scannerNames[0]));
                    menuScanners.getItems().add(checkMenuItemScanner0);
                    checkMenuItemScanner0.setOnAction((event -> {
                        if(checkMenuItemScanner0.isSelected()){
                            connectScanner(0);
                        }
                    }));
                    break;
                case (1):
                    CheckMenuItem checkMenuItemScanner1 = new CheckMenuItem(scannerNames[1]);
                    menuScanners.getItems().add(checkMenuItemScanner1);
                    checkMenuItemScanner1.setOnAction((event -> {
                        if(checkMenuItemScanner1.isSelected()){
                            connectScanner(1);
                        }
                    }));
                    break;

                case (2):
                    CheckMenuItem checkMenuItemScanner2 = new CheckMenuItem(scannerNames[2]);
                    menuScanners.getItems().add(checkMenuItemScanner2);
                    checkMenuItemScanner2.setOnAction((event -> {
                        if(checkMenuItemScanner2.isSelected()){
                            connectScanner(2);
                        }
                    }));
                    break;

            }
        }

    }

    public void clickMenuItemSettings(){
        try {
            SetParams.settings(paramsWorkbook,EXCITATION_WAVELENGTH, TARGET_WAVELENGTH, BOXCAR_WIDTH, INTEGRATION_TIME, METRIC, FLIGHTLINE_LENGTH, SAMPLE_LENGTH, STRING_VEL, isFlipPattern, useLogo, logoFile, invertPH, serialPort, useSpreadFactor, spreadA, spreadB, spreadC, DEFAULT_THRESHOLD, CROP_SCALE, CARD_SPACING);
        } catch (Exception e) {
            e.printStackTrace();
        }
        resetHorizontalAxes(FLIGHTLINE_LENGTH);
        //ToDo Commented for testing, maybe leave commented? It should happen anytime RUN is called
        //setupSpectrometer();
        updateLabelUnits();
        setCardSpacing();
        updateParamsWorkbook();
    }

    //Resources Menu
    public void clickMenuItemUserManual() throws IOException {
        if (Desktop.isDesktopSupported()) {
            try {
                File myFile = new File(userManual);
                //File myFile = new File(getClass().getResource(userManual).getFile());
                Desktop.getDesktop().open(myFile);
            } catch (IOException ex) {
                // no application registered for PDFs
            }

        }


    }

    public void clickMenuItemCalibrationCalc(){
        CalibrationCalc.display();
    }

    public void clickMenuItemAbout(){
        //PatternObject testerObj = new PatternObject();
        //testerObj.setNumPasses(3);
        //System.out.println("Num of Passes" +testerObj.getNumPasses());


        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle("About AccuPatt "+version);
        alert.setHeaderText("Continuing to be developed by:");

        alert.setContentText("Matt Gill\n\nIn consultation with:\n\tDr. Scott Bretthauer\nand in consultation with and marketed by:\n\tDr. Richard Whitney (WRK of Oklahoma)\n\tDr. Dennis Gardisser (WRK of Arkansas)");
        alert.showAndWait();
    }

    //ScrollPane
    public void initializeComboBoxes(){
        //Setup State Combo Box
        textFieldState.clear();

        //Set up Aircraft Make ComboBox
        List aircraftMakes = new ArrayList();
        try{
            FileInputStream fis = new FileInputStream(aircraftWorkbook);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            int numOfSheets = wb.getNumberOfSheets();
            for(int i=0; i<numOfSheets; i++){
                aircraftMakes.add(wb.getSheetAt(i).getSheetName());
            }
            fis.close();
            ObservableList<String> acftMakes = FXCollections.observableArrayList(aircraftMakes);
            comboBoxAircraftMake.getItems().clear();
            comboBoxAircraftMake.setItems(acftMakes);
        } catch (Exception e){
            e.printStackTrace();
        }
        comboBoxAircraftMake.setEditable(true);
        comboBoxAircraftMake.getEditor().textProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                comboBoxAircraftMake.setValue(newValue);
                changeAircraftMake();
            }
        });
        comboBoxAircraftModel.setEditable(true);
        comboBoxAircraftModel.getEditor().textProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                comboBoxAircraftModel.setValue(newValue);
                changeAircraftModel();
            }
        });

        //Set up Nozzle #1 Type Combo Box
        List nozzle1TypeData = new ArrayList();

        try{
            FileInputStream fis = new FileInputStream(nozzleWorkbook);
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
        comboBoxNozzle1Def.setEditable(true);
        comboBoxNozzle1Def.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle1Def.setValue(comboBoxNozzle1Def.getEditor().getText());
        });

        //Set up Nozzle #2 Type Combo Box
        List nozzle2TypeData = new ArrayList();
        try{
            FileInputStream fis = new FileInputStream(nozzleWorkbook);
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
        comboBoxNozzle2Def.setEditable(true);
        comboBoxNozzle2Def.getEditor().textProperty().addListener(e -> {
            comboBoxNozzle2Def.setValue(comboBoxNozzle2Def.getEditor().getText());
        });

        //Set Defaults for Size/Def Boxes
        comboBoxNozzle1Size.setValue("N/A");
        comboBoxNozzle1Def.setValue("N/A");
        comboBoxNozzle2Size.setValue("N/A");
        comboBoxNozzle2Def.setValue("N/A");

        //Extra Info
        ObservableList<String> winglets = FXCollections.observableArrayList("Yes", "No");
        comboBoxWinglets.getItems().clear();
        comboBoxWinglets.setItems(winglets);
        comboBoxWinglets.setValue("No");
    }

    public void changeAircraftMake(){
        //Set up Aircraft Model ComboBox
        List aircraftModelData = new ArrayList();
        comboBoxAircraftModel.getItems().clear();
        try {
            FileInputStream fis = new FileInputStream(aircraftWorkbook);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            if (wb.getSheet(comboBoxAircraftMake.getSelectionModel().getSelectedItem().toString()) != null) {
                XSSFSheet sh = wb.getSheet(comboBoxAircraftMake.getSelectionModel().getSelectedItem().toString());
                Iterator rows = sh.rowIterator();
                while (rows.hasNext()) {
                    XSSFRow row = (XSSFRow) rows.next();
                    Iterator cells = row.cellIterator();
                    List cellData = new ArrayList();
                    while (cells.hasNext()) {
                        XSSFCell cell = (XSSFCell) cells.next();
                        int columnIndex = cell.getColumnIndex();
                        int rowIndex = cell.getRowIndex();
                        if (columnIndex == 0 && rowIndex != 0) {
                            switch (cell.getCellType()) {
                                case XSSFCell.CELL_TYPE_STRING:
                                    cellData.add(cell.getStringCellValue());
                                    break;
                                case XSSFCell.CELL_TYPE_NUMERIC:
                                    cellData.add(Double.toString(cell.getNumericCellValue()));
                                    break;
                            }
                        }
                    }
                    aircraftModelData.addAll(cellData);
                }
                ObservableList<String> aircraftModels = FXCollections.observableArrayList(aircraftModelData);
                comboBoxAircraftModel.getItems().clear();
                comboBoxAircraftModel.setItems(aircraftModels);
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        textFieldWingSpan.clear();
    }

    public void changeAircraftModel(){
        try {
            textFieldWingSpan.clear();
            FileInputStream fis = new FileInputStream(aircraftWorkbook);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            if(wb.getSheet(comboBoxAircraftMake.getValue().toString()) != null){
                if(wb.getSheet(comboBoxAircraftMake.getValue().toString()).getRow(
                        comboBoxAircraftModel.getSelectionModel().getSelectedIndex()) != null){
                    XSSFSheet sh = wb.getSheet(comboBoxAircraftMake.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row = sh.getRow(comboBoxAircraftModel.getSelectionModel().getSelectedIndex()+1);
                    if (row.getCell(0).getStringCellValue().equals(comboBoxAircraftModel.getValue().toString())) {
                        if (row.getCell(2).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                            if (!METRIC) {
                                textFieldWingSpan.setText(Double.toString(row.getCell(2).getNumericCellValue()));
                            } else {
                                textFieldWingSpan.setText(String.format("%.2f",row.getCell(2).getNumericCellValue()*0.3048));
                            }
                        }
                    }
                }
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void changeNozzle1Type() {
        //Set up Nozzle #1 Size Combo Box
        List nozzle1SizeData = new ArrayList();

        //New Method
        if (comboBoxNozzle1Type.getValue() != "Select Nozzle") {
            try{
                FileInputStream fis = new FileInputStream(nozzleWorkbook);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                List cellData = null;
                if (wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString())!=null) {
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row3 = sh.getRow(3);
                    Iterator cells = row3.cellIterator();
                    cellData = new ArrayList();
                    while (cells.hasNext()) {
                        XSSFCell cell = (XSSFCell) cells.next();
                        if (cell.getColumnIndex() != 0) {
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
                    ObservableList<String> nozzle1Sizes = FXCollections.observableArrayList(nozzle1SizeData);
                    comboBoxNozzle1Size.getItems().clear();
                    comboBoxNozzle1Size.setItems(nozzle1Sizes);
                }
                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
        }

        //Set up Nozzle #1 Def Combo Box
        List nozzle1DefData = new ArrayList();

        if (comboBoxNozzle1Type.getValue() != "Select Nozzle") {
            try{
                FileInputStream fis = new FileInputStream(nozzleWorkbook);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                List cellData = null;
                if (wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString())!=null) {
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row4 = sh.getRow(4);
                    Iterator cells = row4.cellIterator();
                    cellData = new ArrayList();
                    while (cells.hasNext()) {
                        XSSFCell cell = (XSSFCell) cells.next();
                        if (cell.getColumnIndex() != 0) {
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
                    nozzle1DefData.addAll(cellData);
                    ObservableList<String> nozzle1Defs = FXCollections.observableArrayList(nozzle1DefData);
                    comboBoxNozzle1Def.getItems().clear();
                    comboBoxNozzle1Def.setItems(nozzle1Defs);
                }
                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    public void changeNozzle2Type() {
        //Set up Nozzle #2 Size Combo Box
        List nozzle2SizeData = new ArrayList();

        if (comboBoxNozzle2Type.getValue() != "Select Nozzle") {
            try{
                FileInputStream fis = new FileInputStream(nozzleWorkbook);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                List cellData = null;
                if (wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString())!=null) {
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row3 = sh.getRow(3);
                    Iterator cells = row3.cellIterator();
                    cellData = new ArrayList();
                    while (cells.hasNext()) {
                        XSSFCell cell = (XSSFCell) cells.next();
                        if (cell.getColumnIndex() != 0) {
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
                    ObservableList<String> nozzle2Sizes = FXCollections.observableArrayList(nozzle2SizeData);
                    comboBoxNozzle2Size.getItems().clear();
                    comboBoxNozzle2Size.setItems(nozzle2Sizes);
                }
                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
        }

        //Set up Nozzle #2 Def Combo Box
        List nozzle2DefData = new ArrayList();

        if (comboBoxNozzle2Type.getValue() != "Select Nozzle") {
            try{
                FileInputStream fis = new FileInputStream(nozzleWorkbook);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                List cellData = null;
                if (wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString())!=null) {
                    XSSFSheet sh = wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString());
                    XSSFRow row4 = sh.getRow(4);
                    Iterator cells = row4.cellIterator();
                    cellData = new ArrayList();
                    while (cells.hasNext()) {
                        XSSFCell cell = (XSSFCell) cells.next();
                        if (cell.getColumnIndex() != 0) {
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
                    nozzle2DefData.addAll(cellData);
                    ObservableList<String> nozzle2Defs = FXCollections.observableArrayList(nozzle2DefData);
                    comboBoxNozzle2Def.getItems().clear();
                    comboBoxNozzle2Def.setItems(nozzle2Defs);
                }
                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    public void updateLabelUnits(){
        labelBoomPress.setText("Boom Pressure ("+uPressure+"): ");
        labelTargetRate.setText("Target Rate ("+uAppRate+"): ");
        labelTargetSwath.setText("Target Swath ("+uHeight+"): ");
        labelGroundSpeed.setText("Ground Speed ("+uAirSpeed+"): ");
        labelSprayHeight.setText("Spray Height ("+uHeight+"): ");
        labelWindVel.setText("Wind Velocity ("+uWindSpeed+"): ");
        labelTemp.setText("Ambient Temp ("+uTemp+"): ");
        labelWingSpan.setText("Wing Span ("+uHeight+"): ");
        labelBoomWidth.setText("Boom Width ("+uHeight+"): ");
        labelBoomDrop.setText("Boom Drop ("+uHeightSmall+"): ");
        labelNozzleSpacing.setText("Nozzle Spacing ("+uHeightSmall+"): ");
        labelSwathUnits.setText(" "+ uHeight);
        labelSwathUnitsR.setText(uHeight);
        labelSwathUnitsBF.setText(uHeight);
    }

    //Individual Passes Tab Controls
    public void clickButtonManualReverse(){
        if(!isReversing) {
            try {
                //Start Reversing onto Supply
                serialPort.writeString("BD-\r");
            } catch (SerialPortException e) {
                e.printStackTrace();
            }
            //Change Text on Reverse Button
            labelManualReverse.setText("Stop");

            //Disable Other Buttons While Reversing
            buttonManualAdvance.setDisable(true);
            buttonManualReverse.setDisable(false);

            isReversing = true;
        }
        else if (isReversing) {
            try{
                //Stop Reversing Takeup
                serialPort.writeString("BD\r");
            } catch (SerialPortException e) {
                e.printStackTrace();
            }
            //Change Text on Reverse Button
            labelManualReverse.setText("Start");

            //Re-Enable Other Buttons
            buttonManualAdvance.setDisable(false);
            buttonManualReverse.setDisable(false);
            isReversing = false;
        }
    }

    public void clickButtonManualAdvance(){
        if(!isAdvancing) {
            try {
                //Start Advancing
                serialPort.writeString("AD+\r");
            } catch (SerialPortException e) {
                e.printStackTrace();
            }
            //Update Label on Advance Button
            labelManualAdvance.setText("Stop");

            //Disable Other Buttons While Advancing
            buttonManualReverse.setDisable(true);
            isAdvancing = true;
        }
        else if (isAdvancing) {
            try{
                //Stop Advancing
                serialPort.writeString("AD\r");
            } catch (SerialPortException e) {
                e.printStackTrace();
            }
            //Update Label on Advance Button
            labelManualAdvance.setText("Start");

            //Re-Enable Other Buttons
            buttonManualReverse.setDisable(false);
            buttonManualReverse.setDisable(false);
            isAdvancing = false;
        }
    }

    public void clickButtonPassStart(){
        double fl = FLIGHTLINE_LENGTH;
        double sl = SAMPLE_LENGTH;

        if(spinnerPassNum.getValue().equals(1) && !isScannedPass1){
            if (!pass1Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass1Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 1);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();

                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 1
                animation1 = new Timeline();
                animation1.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));
                animation1.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation1.play();
                animation1.setOnFinished(event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}
                    isScannedPass1 = true;
                    toggleButtonPass1.setDisable(false);
                    toggleButtonPass1.setSelected(true);
                    updateVisibilityAnalyzing(false, 1);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass1(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(1).setCellValue(fl);
                        sh.getRow(2).createCell(1).setCellValue(sl);
                        sh.getRow(3).createCell(1).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(1).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(2).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(1).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(2).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(1);
                        sh.autoSizeColumn(2);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                });
            }
        } else if (spinnerPassNum.getValue().equals(2) && !isScannedPass2){
            buttonPassStop.setText("Stop");
            if (!pass2Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass2Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 2);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();
                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 2
                animation2 = new Timeline();
                animation2.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));
                animation2.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation2.play();
                animation2.setOnFinished((event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}
                    isScannedPass2 = true;
                    toggleButtonPass2.setDisable(false);
                    toggleButtonPass2.setSelected(true);
                    updateVisibilityAnalyzing(false, 2);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass2(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(3).setCellValue(fl);
                        sh.getRow(2).createCell(3).setCellValue(sl);
                        sh.getRow(3).createCell(3).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(3).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(4).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(3).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(4).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(3);
                        sh.autoSizeColumn(4);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                }));
            }
        } else if (spinnerPassNum.getValue().equals(3) && !isScannedPass3){
            buttonPassStop.setText("Stop");
            if (!pass3Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass3Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 3);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();
                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 3
                animation3 = new Timeline();
                animation3.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));

                animation3.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation3.play();
                animation3.setOnFinished((event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}

                    isScannedPass3 = true;
                    toggleButtonPass3.setDisable(false);
                    toggleButtonPass3.setSelected(true);
                    updateVisibilityAnalyzing(false, 3);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass3(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(5).setCellValue(fl);
                        sh.getRow(2).createCell(5).setCellValue(sl);
                        sh.getRow(3).createCell(5).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(5).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(6).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(5).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(6).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(5);
                        sh.autoSizeColumn(6);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                }));
            }
        } else if (spinnerPassNum.getValue().equals(4) && !isScannedPass4){
            buttonPassStop.setText("Stop");
            if (!pass4Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass4Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 4);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();
                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 3
                animation4 = new Timeline();
                animation4.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));

                animation4.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation4.play();
                animation4.setOnFinished((event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}

                    isScannedPass4 = true;
                    toggleButtonPass4.setDisable(false);
                    toggleButtonPass4.setSelected(true);
                    updateVisibilityAnalyzing(false, 4);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass4(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(7).setCellValue(fl);
                        sh.getRow(2).createCell(7).setCellValue(sl);
                        sh.getRow(3).createCell(7).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(7).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(8).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(7).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(8).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(7);
                        sh.autoSizeColumn(8);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                }));
            }
        } else if (spinnerPassNum.getValue().equals(5) && !isScannedPass5){
            buttonPassStop.setText("Stop");
            if (!pass5Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass5Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 5);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();
                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 3
                animation5 = new Timeline();
                animation5.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));

                animation5.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation5.play();
                animation5.setOnFinished((event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}

                    isScannedPass5 = true;
                    toggleButtonPass5.setDisable(false);
                    toggleButtonPass5.setSelected(true);
                    updateVisibilityAnalyzing(false, 5);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass5(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(9).setCellValue(fl);
                        sh.getRow(2).createCell(9).setCellValue(sl);
                        sh.getRow(3).createCell(9).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(9).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(10).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(9).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(10).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(9);
                        sh.autoSizeColumn(10);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                }));
            }
        } else if (spinnerPassNum.getValue().equals(6) && !isScannedPass6){
            buttonPassStop.setText("Stop");
            if (!pass6Mark) {
                try {
                    serialPort.writeString("AD+\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                buttonPassStart.setText("Mark");
                pass6Mark = true;
            } else {
                buttonPassStart.setText("Start");
                updateVisibilityAnalyzing(true, 6);
                ArrayList<Double> passExcitation = new ArrayList<>();
                ArrayList<Double> passPattern = new ArrayList<>();
                AreaChart.Series passSeries = new AreaChart.Series();
                location = -(fl/2);
                cycleNum = 0;
                //Declare then Define the animation for Pass 3
                animation6 = new Timeline();
                animation6.getKeyFrames().add(new KeyFrame(Duration.seconds((sl*STRING_TIME)/fl), (event -> {
                    //GetSpectrum
                    double [] run = wrapper.getSpectrum(0);
                    //Add fluorescence value
                    passPattern.add(run[TARGET_PIXEL]);
                    //And Excitation value for fun
                    passExcitation.add(run[EXCITATION_PIXEL]);
                    //Add to Plotting Series for Live Updating
                    passSeries.getData().add(new AreaChart.Data<Number, Number>(location,run[TARGET_PIXEL]));
                    //Increment
                    location += (sl);
                    cycleNum++;
                })));

                animation6.setCycleCount((int) (fl/sl)+1);
                areaChartPass.getData().addAll(passSeries);
                animation6.play();
                animation6.setOnFinished((event ->{
                    try {
                        //Stop Advancing String
                        serialPort.writeString("AD\r");
                    } catch (SerialPortException e) {e.printStackTrace();}

                    isScannedPass6 = true;
                    toggleButtonPass6.setDisable(false);
                    toggleButtonPass6.setSelected(true);
                    updateVisibilityAnalyzing(false, 6);

                    patternObject.setSampleLength(sl);
                    patternObject.setPass6(passPattern);
                    refreshPlots1();
                    refreshPlots2();

                    //Save Pattern Data to Current Data File
                    try {
                        FileInputStream fis = new FileInputStream(currentDataFile);
                        Workbook wb = WorkbookFactory.create(fis);
                        Sheet sh = wb.getSheet("Pattern Data");
                        sh.getRow(1).createCell(11).setCellValue(fl);
                        sh.getRow(2).createCell(11).setCellValue(sl);
                        sh.getRow(3).createCell(11).setCellValue(INTEGRATION_TIME);
                        sh.getRow(4).createCell(11).setCellValue(round(wrapper.getWavelength(0,EXCITATION_PIXEL),1)+"nm ("+EXCITATION_PIXEL+")");
                        sh.getRow(4).createCell(12).setCellValue(round(wrapper.getWavelength(0,TARGET_PIXEL),1)+"nm ("+TARGET_PIXEL+")");
                        for (int i = 0; i < passPattern.size(); i++) {
                            sh.getRow(i+5).createCell(11).setCellValue(passExcitation.get(i));
                            sh.getRow(i+5).createCell(12).setCellValue(passPattern.get(i));
                        }
                        sh.autoSizeColumn(11);
                        sh.autoSizeColumn(12);
                        FileOutputStream fos = new FileOutputStream(currentDataFile);
                        wb.write(fos);
                        fos.close();
                    } catch (Exception e) {e.printStackTrace();}
                }));
            }
        }
    }

    public void clickButtonPassStop(){
        if (spinnerPassNum.getValue().equals(1)) {
            if(isScannedPass1){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 1?");
                alert.setContentText("All data from pass 1 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    pass1Mark = false;
                    isScannedPass1 = false;
                    updateVisibilityAnalyzing(false, 1);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass1Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation1.stop();
                    animation1.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 1);
                }
            }
        } else if (spinnerPassNum.getValue().equals(2)){
            if(isScannedPass2){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 2?");
                alert.setContentText("All data from pass 2 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    isScannedPass2 = false;
                    pass2Mark = false;
                    updateVisibilityAnalyzing(false, 2);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass2Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation2.stop();
                    animation2.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 2);
                }
            }
        } else if (spinnerPassNum.getValue().equals(3)){
            if(isScannedPass3){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 3?");
                alert.setContentText("All data from pass 3 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    isScannedPass3 = false;
                    pass3Mark = false;
                    updateVisibilityAnalyzing(false, 3);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass3Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation3.stop();
                    animation3.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 3);
                }
            }
        } else if (spinnerPassNum.getValue().equals(4)){
            if(isScannedPass4){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 4?");
                alert.setContentText("All data from pass 4 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    isScannedPass4 = false;
                    pass4Mark = false;
                    updateVisibilityAnalyzing(false, 4);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass4Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation4.stop();
                    animation4.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 4);
                }
            }
        } else if (spinnerPassNum.getValue().equals(5)){
            if(isScannedPass5){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 5?");
                alert.setContentText("All data from pass 5 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    isScannedPass5 = false;
                    pass5Mark = false;
                    updateVisibilityAnalyzing(false, 5);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass5Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation5.stop();
                    animation5.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 5);
                }
            }
        } else if (spinnerPassNum.getValue().equals(6)){
            if(isScannedPass6){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Repeat Pass 6?");
                alert.setContentText("All data from pass 6 will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    areaChartPass.getData().clear();
                    areaChartPassTrim.getData().clear();
                    isScannedPass6 = false;
                    pass6Mark = false;
                    updateVisibilityAnalyzing(false, 6);
                }
            } else {
                try {
                    serialPort.writeString("AD\r");
                } catch (SerialPortException e) {
                    e.printStackTrace();
                }
                pass6Mark = false;
                if(buttonPassStart.getText().equals("Mark")){
                    buttonPassStart.setText("Start");
                } else {
                    animation6.stop();
                    animation6.getKeyFrames().clear();
                    areaChartPass.getData().clear();
                    updateVisibilityAnalyzing(false, 6);
                }
            }
        }
    }

    //Average Tab Controls
    public void clickToggleButtonAlignCentroid(){
        isAlignCentroid = toggleButtonAlignCentroid.isSelected();
        if(isAlignCentroid){
            toggleButtonAlignCentroid.setText("Align Centroid: ON");
        } else {
            toggleButtonAlignCentroid.setText("Align Centroid: OFF");
        }
        refreshPlots1();
        refreshPlots2();
    }

    public void clickToggleButtonSmoothSeries(){
        isSmoothSeries = toggleButtonSmoothSeries.isSelected();
        if(isSmoothSeries){
            toggleButtonSmoothSeries.setText("Smooth Series: ON");
        } else {
            toggleButtonSmoothSeries.setText("Smooth Series: OFF");
        }
        refreshPlots1();
        refreshPlots2();
    }

    public void clickToggleButtonPass1(){
        isScannedPass1 = toggleButtonPass1.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonPass2(){
        isScannedPass2 = toggleButtonPass2.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonPass3(){
        isScannedPass3 = toggleButtonPass3.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonPass4(){
        isScannedPass4 = toggleButtonPass4.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonPass5(){
        isScannedPass5 = toggleButtonPass5.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonPass6(){
        isScannedPass6 = toggleButtonPass6.isSelected();
        refreshPlots1();
        refreshPlots2();
        updatePassColors();
    }

    public void clickToggleButtonSmoothAverage(){
        isSmoothAverage = toggleButtonSmoothAverage.isSelected();
        if(isSmoothAverage){
            toggleButtonSmoothAverage.setText("Smooth Average: ON");
        } else {
            toggleButtonSmoothAverage.setText("Smooth Average: OFF");
        }
        refreshPlots2();
    }

    public void clickToggleButtonEqualizeArea(){
        isEqualizeArea = toggleButtonEqualizeArea.isSelected();
        if(isEqualizeArea){
            toggleButtonEqualizeArea.setText("Equalize Area: ON");
        } else {
            toggleButtonEqualizeArea.setText("Equalize Area: OFF");
        }
        refreshPlots1();
        refreshPlots2();
    }

    //Swath Analysis Tab Controls
    public void clickButtonSwathDEC(){
        if (!METRIC) {
            labelSwathFinal.setText(Integer.toString(Integer.parseInt(labelSwathFinal.getText())-1));
        } else {
            labelSwathFinal.setText(Double.toString(Double.parseDouble(labelSwathFinal.getText())-0.5));
        }
        refreshPlots2();
        plotRTCV();
        plotBFCV();
    }

    public void clickButtonSwathINC(){
        if (!METRIC) {
            labelSwathFinal.setText(Integer.toString(Integer.parseInt(labelSwathFinal.getText())+1));
        } else {
            labelSwathFinal.setText(Double.toString(Double.parseDouble(labelSwathFinal.getText())+0.5));
        }
        refreshPlots2();
        plotRTCV();
        plotBFCV();
    }

    public void clickToggleButtonLR1(){
        if(cV3){
            toggleButtonL1.setSelected(false);
            toggleButtonR1.setSelected(false);
            cV3 = false;
            cV5 = false;
            plotRTCV();
            plotBFCV();
            toggleButtonL2.setSelected(false);
            toggleButtonR2.setSelected(false);
            toggleButtonL2.setDisable(true);
            toggleButtonR2.setDisable(true);

        } else {
            toggleButtonL1.setSelected(true);
            toggleButtonR1.setSelected(true);
            cV3 = true;
            plotRTCV();
            plotBFCV();
            toggleButtonL2.setDisable(false);
            toggleButtonR2.setDisable(false);
        }
    }

    public void clickToggleButtonLR2(){
        if(cV5){
            toggleButtonL2.setSelected(false);
            toggleButtonR2.setSelected(false);
            cV5 = false;
            plotRTCV();
            plotBFCV();

        } else {
            toggleButtonL2.setSelected(true);
            toggleButtonR2.setSelected(true);
            cV5 = true;
            plotRTCV();
            plotBFCV();
        }
    }

    //File and Report Methods
    private void openAccuPatt(File datafile){
        //Alter Perspective
        scrollPaneMain.setDisable(false);
        buttonRun.setDisable(false);
        buttonRun.setText("Update Data File");
        tabPaneMain.setDisable(false);
        tabPaneMain.getSelectionModel().select(0);
        clearAircraftAndSeriesData();
        clearPatternData();
        //Set Smoothing On
        toggleButtonSmoothSeries.setSelected(true);
        toggleButtonSmoothSeries.setText("Smooth Series: On");
        isSmoothSeries = true;
        toggleButtonSmoothAverage.setSelected(true);
        toggleButtonSmoothAverage.setText("Smooth Average: On");
        isSmoothAverage = true;

        //Retrieve a past AccuPatt file
        FileChooser fc = new FileChooser();
        if (getCurrentDirectory()!=null) {
            fc.setInitialDirectory(getCurrentDirectory());
        }
        fc.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("AccuPatt Files", "*.xlsx"));
        if (datafile != null) {
            if (datafile.getParentFile().exists() && datafile.getParentFile().isDirectory()) {
                currentDataFile = datafile;
                currentDirectory = currentDataFile.getParentFile();
            }
            try {
                FileInputStream fis = new FileInputStream(datafile);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                fis.close();
                //Get Headers
                XSSFSheet sh = wb.getSheet("Fly-In Data");
                if (sh.getRow(0) != null && sh.getRow(0).getCell(0) != null) {
                    header1 = sh.getRow(0).getCell(0).getStringCellValue();
                }
                if (sh.getRow(1) != null && sh.getRow(1).getCell(0) != null) {
                    header2 = sh.getRow(1).getCell(0).getStringCellValue();
                }
                if (sh.getRow(2) != null && sh.getRow(2).getCell(0) != null) {
                    header3 = sh.getRow(2).getCell(0).getStringCellValue();
                }
                if (sh.getRow(3) != null && sh.getRow(3).getCell(0) != null) {
                    header4 = sh.getRow(3).getCell(0).getStringCellValue();
                }

                //Get Aircraft Data
                XSSFSheet sheet = wb.getSheet("Aircraft Data");
                //Determine if metric or not
                if (sheet.getRow(35).getCell(1) != null) {
                    setMetric(sheet.getRow(35).getCell(1).getBooleanCellValue());
                    updateLabelUnits();
                } else {
                    METRIC = false;
                }

                if (sheet.getRow(0).getCell(1) != null) {
                    textFieldPilot.setText(sheet.getRow(0).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(0).getCell(2) != null) {
                    textFieldEmail.setText(sheet.getRow(0).getCell(2).getStringCellValue());
                }
                if (sheet.getRow(1).getCell(1) != null) {
                    textFieldBusinessName.setText(sheet.getRow(1).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(2).getCell(1) != null) {
                    textFieldPhone.setText(sheet.getRow(2).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(3).getCell(1) != null) {
                    textFieldStreet.setText(sheet.getRow(3).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(4).getCell(1) != null) {
                    textFieldCity.setText(sheet.getRow(4).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(5).getCell(1) != null) {
                    textFieldState.setText(sheet.getRow(5).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(5).getCell(2) != null) {
                    textFieldZIP.setText(sheet.getRow(5).getCell(2).getStringCellValue());
                }
                if (sheet.getRow(6).getCell(1) != null) {
                    textFieldRegNum.setText(sheet.getRow(6).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(7).getCell(1) != null) {
                    textFieldSeriesNum.setText(sheet.getRow(7).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(8).getCell(1) != null) {
                    comboBoxAircraftMake.setValue(sheet.getRow(8).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(9).getCell(1) != null) {
                    comboBoxAircraftModel.setValue(sheet.getRow(9).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(10).getCell(1) != null) {
                    comboBoxNozzle1Type.setValue(sheet.getRow(10).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(11).getCell(1) != null) {
                    textFieldNozzle1Quant.setText(sheet.getRow(11).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(12).getCell(1) != null) {
                    comboBoxNozzle1Size.setValue(sheet.getRow(12).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(13).getCell(1) != null) {
                    comboBoxNozzle1Def.setValue(sheet.getRow(13).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(14).getCell(1) != null) {
                    comboBoxNozzle2Type.setValue(sheet.getRow(14).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(15).getCell(1) != null) {
                    textFieldNozzle2Quant.setText(sheet.getRow(15).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(16).getCell(1) != null) {
                    comboBoxNozzle2Size.setValue(sheet.getRow(16).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(17).getCell(1) != null) {
                    comboBoxNozzle2Def.setValue(sheet.getRow(17).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(18).getCell(1) != null) {
                    textFieldBoomPressure.setText(sheet.getRow(18).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(19).getCell(1) != null) {
                    textFieldTargetRate.setText(sheet.getRow(19).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(20).getCell(1) != null) {
                    textFieldTargetSwath.setText(sheet.getRow(20).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(20).getCell(2) != null) {
                    labelSwathFinal.setText(sheet.getRow(20).getCell(2).getStringCellValue());
                } else if(sheet.getRow(20).getCell(1) != null){
                    labelSwathFinal.setText(sheet.getRow(20).getCell(1).getStringCellValue());
                } else {
                    if(!METRIC){
                        labelSwathFinal.setText("50");
                    } else {
                        labelSwathFinal.setText("5");
                    }
                }
                if (sheet.getRow(26).getCell(1) != null) {
                    textFieldTime.setText(sheet.getRow(26).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(27) != null && sheet.getRow(27).getCell(1) != null) {
                    textFieldWingSpan.setText(sheet.getRow(27).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(28) != null && sheet.getRow(28).getCell(1) != null) {
                    textFieldBoomWidth.setText(sheet.getRow(28).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(30) != null && sheet.getRow(30).getCell(1) != null) {
                    textFieldBoomDrop.setText(sheet.getRow(30).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(31) != null && sheet.getRow(31).getCell(1) != null) {
                    textFieldNozzleSpacing.setText(sheet.getRow(31).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(32) != null && sheet.getRow(32).getCell(1) != null) {
                    comboBoxWinglets.setValue(sheet.getRow(32).getCell(1).getStringCellValue());
                }
                if (sheet.getRow(33) != null && sheet.getRow(33).getCell(1) != null) {
                    textAreaNotes.setText(sheet.getRow(33).getCell(1).getStringCellValue());
                }
                //Get Series Data
                XSSFSheet sheet1 = wb.getSheet("Series Data");
                if (sheet1.getRow(1).getCell(1) != null) {
                    textFieldGS1.setText(sheet1.getRow(1).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(1).getCell(2) != null) {
                    textFieldGS2.setText(sheet1.getRow(1).getCell(2).getStringCellValue());
                }
                if (sheet1.getRow(1).getCell(3) != null) {
                    textFieldGS3.setText(sheet1.getRow(1).getCell(3).getStringCellValue());
                }
                if (sheet1.getRow(1).getCell(4) != null) {
                    textFieldGS4.setText(sheet1.getRow(1).getCell(4).getStringCellValue());
                }
                if (sheet1.getRow(1).getCell(5) != null) {
                    textFieldGS5.setText(sheet1.getRow(1).getCell(5).getStringCellValue());
                }
                if (sheet1.getRow(1).getCell(6) != null) {
                    textFieldGS6.setText(sheet1.getRow(1).getCell(6).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(1) != null) {
                    textFieldSH1.setText(sheet1.getRow(2).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(2) != null) {
                    textFieldSH2.setText(sheet1.getRow(2).getCell(2).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(3) != null) {
                    textFieldSH3.setText(sheet1.getRow(2).getCell(3).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(4) != null) {
                    textFieldSH4.setText(sheet1.getRow(2).getCell(4).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(5) != null) {
                    textFieldSH5.setText(sheet1.getRow(2).getCell(5).getStringCellValue());
                }
                if (sheet1.getRow(2).getCell(6) != null) {
                    textFieldSH6.setText(sheet1.getRow(2).getCell(6).getStringCellValue());
                }
                if (sheet1.getRow(3).getCell(0) != null) {
                    if(sheet1.getRow(3).getCell(0).getStringCellValue().contains("*")){
                        invertPH = true;
                    } else {
                        invertPH = false;
                    }
                }
                if (sheet1.getRow(3).getCell(1) != null) {
                    textFieldPH1.setText(sheet1.getRow(3).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(3).getCell(2) != null) {
                    textFieldPH2.setText(sheet1.getRow(3).getCell(2).getStringCellValue());
                }
                if (sheet1.getRow(3).getCell(3) != null) {
                    textFieldPH3.setText(sheet1.getRow(3).getCell(3).getStringCellValue());
                }
                if (sheet1.getRow(3).getCell(4) != null) {
                    textFieldPH4.setText(sheet1.getRow(3).getCell(4).getStringCellValue());
                }
                if (sheet1.getRow(3).getCell(5) != null) {
                    textFieldPH5.setText(sheet1.getRow(3).getCell(5).getStringCellValue());
                }
                if (sheet1.getRow(6).getCell(3) != null) {
                    textFieldPH6.setText(sheet1.getRow(3).getCell(6).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(1) != null) {
                    textFieldWD1.setText(sheet1.getRow(4).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(2) != null) {
                    textFieldWD2.setText(sheet1.getRow(4).getCell(2).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(3) != null) {
                    textFieldWD3.setText(sheet1.getRow(4).getCell(3).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(4) != null) {
                    textFieldWD4.setText(sheet1.getRow(4).getCell(4).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(5) != null) {
                    textFieldWD5.setText(sheet1.getRow(4).getCell(5).getStringCellValue());
                }
                if (sheet1.getRow(4).getCell(6) != null) {
                    textFieldWD6.setText(sheet1.getRow(4).getCell(6).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(1) != null) {
                    textFieldWV1.setText(sheet1.getRow(5).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(2) != null) {
                    textFieldWV2.setText(sheet1.getRow(5).getCell(2).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(3) != null) {
                    textFieldWV3.setText(sheet1.getRow(5).getCell(3).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(4) != null) {
                    textFieldWV4.setText(sheet1.getRow(5).getCell(4).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(5) != null) {
                    textFieldWV5.setText(sheet1.getRow(5).getCell(5).getStringCellValue());
                }
                if (sheet1.getRow(5).getCell(6) != null) {
                    textFieldWV6.setText(sheet1.getRow(5).getCell(6).getStringCellValue());
                }
                if (sheet1.getRow(6).getCell(1) != null) {
                    textFieldAT.setText(sheet1.getRow(6).getCell(1).getStringCellValue());
                }
                if (sheet1.getRow(7).getCell(1) != null) {
                    textFieldRH.setText(sheet1.getRow(7).getCell(1).getStringCellValue());
                }
                //Get Pattern Data
                patternObject = new PatternObject();
                isScannedPass1 = false;
                isScannedPass2 = false;
                isScannedPass3 = false;
                isScannedPass4 = false;
                isScannedPass5 = false;
                isScannedPass6 = false;
                XSSFSheet sheet2 = wb.getSheet("Pattern Data");
                if(sheet2.getRow(5) != null && sheet2.getRow(5).getCell(0) != null) {
                    int firstIndex = 0;
                    for(int i=1; i<=11; i+=2){
                        if(sheet2.getRow(1).getCell(i) != null){
                            firstIndex = i;
                            break;
                        }
                    }
                    setFlightlineLength(sheet2.getRow(1).getCell(firstIndex).getNumericCellValue());
                    setSampleLength(sheet2.getRow(2).getCell(firstIndex).getNumericCellValue());
                    patternObject.setSampleLength(sheet2.getRow(2).getCell(firstIndex).getNumericCellValue());
                    resetHorizontalAxes(FLIGHTLINE_LENGTH);

                    //Create Temp containers for patterns
                    ArrayList<Double> pass1 = new ArrayList<>();
                    ArrayList<Double> pass2 = new ArrayList<>();
                    ArrayList<Double> pass3 = new ArrayList<>();
                    ArrayList<Double> pass4 = new ArrayList<>();
                    ArrayList<Double> pass5 = new ArrayList<>();
                    ArrayList<Double> pass6 = new ArrayList<>();

                    //Get Pattern info if present
                    for(Row row : sheet2){
                        //Get Left Trims
                        if(row.getRowNum()==0) {
                            if (row.getCell(2) != null) {
                                patternObject.setTrim1L((int) row.getCell(2).getNumericCellValue());
                            }
                            if (row.getCell(4) != null) {
                                patternObject.setTrim2L((int) row.getCell(4).getNumericCellValue());
                            }
                            if(row.getCell(6)!=null){
                                patternObject.setTrim3L((int) row.getCell(6).getNumericCellValue());
                            }
                            if (row.getCell(8) != null) {
                                patternObject.setTrim4L((int) row.getCell(8).getNumericCellValue());
                            }
                            if (row.getCell(10) != null) {
                                patternObject.setTrim5L((int) row.getCell(10).getNumericCellValue());
                            }
                            if(row.getCell(12)!=null){
                                patternObject.setTrim6L((int) row.getCell(12).getNumericCellValue());
                            }
                        }
                        //Get Right Trims
                        if(row.getRowNum()==1) {
                            if(row.getCell(2) != null) {
                                patternObject.setTrim1R((int) row.getCell(2).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim1R(patternObject.getTrim1L());
                            }
                            if(row.getCell(4) != null) {
                                patternObject.setTrim2R((int) row.getCell(4).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim2R(patternObject.getTrim2L());
                            }
                            if(row.getCell(6) != null) {
                                patternObject.setTrim3R((int) row.getCell(6).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim3R(patternObject.getTrim3L());
                            }
                            if(row.getCell(8) != null) {
                                patternObject.setTrim4R((int) row.getCell(8).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim4R(patternObject.getTrim4L());
                            }
                            if(row.getCell(10) != null) {
                                patternObject.setTrim5R((int) row.getCell(10).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim5R(patternObject.getTrim5L());
                            }
                            if(row.getCell(12) != null) {
                                patternObject.setTrim6R((int) row.getCell(12).getNumericCellValue());
                            } else {
                                //Backwards Compatible to when only 1 value for horizontal trim existed
                                patternObject.setTrim6R(patternObject.getTrim6L());
                            }
                        }
                        //Get Vertical Trims
                        if(row.getRowNum()==2){
                            if(row.getCell(2)!=null){
                                patternObject.setTrimVertical1(row.getCell(2).getNumericCellValue());
                            }
                            if(row.getCell(4)!=null){
                                patternObject.setTrimVertical2(row.getCell(4).getNumericCellValue());
                            }
                            if(row.getCell(6)!=null){
                                patternObject.setTrimVertical3(row.getCell(6).getNumericCellValue());
                            }
                            if(row.getCell(8)!=null){
                                patternObject.setTrimVertical4(row.getCell(8).getNumericCellValue());
                            }
                            if(row.getCell(10)!=null){
                                patternObject.setTrimVertical5(row.getCell(10).getNumericCellValue());
                            }
                            if(row.getCell(12)!=null){
                                patternObject.setTrimVertical6(row.getCell(12).getNumericCellValue());
                            }
                        }
                        //Get Patterns
                        if(row.getRowNum()>4){
                            if(row.getCell(2)!=null) {
                                pass1.add(row.getCell(2).getNumericCellValue());
                            }
                            if(row.getCell(4)!=null) {
                                pass2.add(row.getCell(4).getNumericCellValue());
                            }
                            if(row.getCell(6)!=null) {
                                pass3.add(row.getCell(6).getNumericCellValue());
                            }
                            if(row.getCell(8)!=null) {
                                pass4.add(row.getCell(8).getNumericCellValue());
                            }
                            if(row.getCell(10)!=null) {
                                pass5.add(row.getCell(10).getNumericCellValue());
                            }
                            if(row.getCell(12)!=null) {
                                pass6.add(row.getCell(12).getNumericCellValue());
                            }
                        }
                    }
                    //Plot present passes
                    if(pass1.size()>0){
                        patternObject.setPass1(pass1);
                        isScannedPass1 = true;
                        cBPass1.setSelected(true);
                        toggleButtonPass1.setDisable(false);
                        toggleButtonPass1.setSelected(true);
                    }
                    if(pass2.size()>0){
                        patternObject.setPass2(pass2);
                        isScannedPass2 = true;
                        cBPass2.setSelected(true);
                        toggleButtonPass2.setDisable(false);
                        toggleButtonPass2.setSelected(true);
                        System.out.println("TEST");
                    }
                    if(pass3.size()>0){
                        patternObject.setPass3(pass3);
                        isScannedPass3 = true;
                        cBPass3.setSelected(true);
                        toggleButtonPass3.setDisable(false);
                        toggleButtonPass3.setSelected(true);
                    }
                    if(pass4.size()>0){
                        patternObject.setPass4(pass4);
                        isScannedPass4 = true;
                        cBPass4.setSelected(true);
                        toggleButtonPass4.setDisable(false);
                        toggleButtonPass4.setSelected(true);
                    }
                    if(pass5.size()>0){
                        patternObject.setPass5(pass5);
                        isScannedPass5 = true;
                        cBPass5.setSelected(true);
                        toggleButtonPass5.setDisable(false);
                        toggleButtonPass5.setSelected(true);
                    }
                    if(pass6.size()>0){
                        patternObject.setPass6(pass6);
                        isScannedPass6 = true;
                        cBPass6.setSelected(true);
                        toggleButtonPass6.setDisable(false);
                        toggleButtonPass6.setSelected(true);
                    }
                    resetHorizontalAxes(FLIGHTLINE_LENGTH);
                    updateActivePasses();
                    if (isScannedPass1 || isScannedPass2 || isScannedPass3 || isScannedPass4 || isScannedPass5 || isScannedPass6) {
                        updatePass();
                    }
                    refreshPlots1();
                    refreshPlots2();
                    updatePassColors();
                }

                if((wb.getSheet("Card Data"))!=null) {
                    openAccuStain(wb);
                } else if((wb.getSheet("Card Index"))!=null) {
                    openAccuStainLegacy(wb);
                }
            } catch (Exception e) {
                e.printStackTrace();
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("File Open Error");
                alert.setHeaderText("Error Opening File");
                alert.setContentText("Check file contents and try again.");
                alert.showAndWait();            }

        }

        labelCurrentFile.setText(currentDataFile.toString());
    }                //XLSX

    private void openWRK(File datafile){
        //Clearing
        scrollPaneMain.setDisable(false);
        tabPaneMain.setDisable(false);
        clearAircraftAndSeriesData();
        clearPatternData();
        //Set smoothing off
        toggleButtonSmoothSeries.setSelected(false);
        toggleButtonSmoothSeries.setText("Smooth Series: Off");
        isSmoothSeries = false;
        toggleButtonSmoothAverage.setSelected(false);
        toggleButtonSmoothAverage.setText("Smooth Average: Off");
        isSmoothAverage = false;
        //Retrieve a past WRK file
        String runNum = "";
        if(datafile.toString().length() >=2){
            runNum = datafile.toString().substring(datafile.toString().length() -2,
                    datafile.toString().length()-1);
        }
        textFieldSeriesNum.setText(runNum);
        //Extract Data from the "A" File
        ArrayList<String> line = new ArrayList<>();
        ArrayList<Double> passPattern = new ArrayList<>();
        int numDataPoints = 0;
        int numPasses = 0;
        double sampleLength = 0;
        try {
            BufferedReader in = new BufferedReader(new FileReader(datafile));

            String lineEntry;
            int q = 0;
            while ((lineEntry = in.readLine()) != null) {
                lineEntry = lineEntry.replace("\"",""); //Remove quotes
                if(q<37){
                   line.add(lineEntry);
                } else if (q==37){
                    numDataPoints = Integer.parseInt(lineEntry);
                } else if (q>37 && q<=37+numDataPoints){
                    passPattern.add(Double.parseDouble(lineEntry));
                } else {
                    numPasses = Integer.parseInt(lineEntry);
                }
                q++;
            }
        } catch (Exception e){
            e.printStackTrace();
        }
        //Decode Data for Pass 1
        header1 = line.get(0);
        header2 = line.get(1);
        header3 = line.get(2);
        header3 = line.get(3);
        //Flightline Length = line.get(4);
        //Analysis Speed = line.get(5);
        //Type 1 Orifice Flow rate at 40 PSI = line.get(6);
        //Type 2 Orifice Flow rate at 40 PSI = line.get(7);
        textFieldNozzle2Quant.setText(line.get(8));
        comboBoxNozzle1Size.setValue(line.get(9));
        comboBoxNozzle2Size.setValue(line.get(10));
        //Type 1 Orifice Size Name = line.get(11);
        //Type 2 Orifice Size Name = line.get(12);
        comboBoxNozzle1Def.setValue(line.get(13));
        comboBoxNozzle2Def.setValue(line.get(14));
        textFieldBusinessName.setText(line.get(15));
        textFieldStreet.setText(line.get(16));
        textFieldCity.setText(line.get(17));
        textFieldState.setText(line.get(18));
        textFieldZIP.setText(line.get(19));
        textFieldPhone.setText(line.get(20));
        textFieldPilot.setText(line.get(21));
        textFieldRegNum.setText(line.get(22));
        comboBoxAircraftModel.setValue(line.get(23));
        comboBoxNozzle1Type.setValue(line.get(24));
        if (Integer.parseInt(line.get(8))>0) {
            comboBoxNozzle2Type.setValue(line.get(24));
            textFieldNozzle1Quant.setText(Integer.toString(Integer.parseInt(line.get(25))-Integer.parseInt(line.get(8))));
        } else {
            textFieldNozzle1Quant.setText(line.get(25));
        }
        textFieldBoomPressure.setText(line.get(26));
        textFieldTargetRate.setText(line.get(27));
        textFieldTargetSwath.setText(line.get(28));
        labelSwathFinal.setText(line.get(28));
        textFieldGS1.setText(line.get(29));
        textFieldWV1.setText(line.get(30));
        textFieldWD1.setText(line.get(31));
        double pHA = 0;
        if(Double.parseDouble(line.get(32)) > 0) {
            pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
        } else {
            pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
        }
        textFieldPH1.setText(Double.toString(pHA));
        textFieldAT.setText(line.get(33));
        textFieldSH1.setText(line.get(34));
        textFieldRH.setText(line.get(35));

        patternObject = new PatternObject();
        patternObject.setSampleLength(Double.parseDouble(line.get(4))/numDataPoints);
        patternObject.setPass1(passPattern);
        cBPass1.setSelected(true);
        isScannedPass1 = true;
        toggleButtonPass1.setDisable(false);
        toggleButtonPass1.setSelected(true);

        //Read in Pass 2
        if(numPasses>1){
            line = new ArrayList<>();
            passPattern = new ArrayList<>();
            try {
                String datafileB = (datafile.toString().substring(0, datafile.toString().length() -1) + "B");
                BufferedReader in = new BufferedReader(new FileReader(datafileB));

                String lineEntry;
                int q = 0;
                while ((lineEntry = in.readLine()) != null) {
                    lineEntry = lineEntry.replace("\"",""); //Remove quotes
                    if(q<37){
                        line.add(lineEntry);
                    } else if (q==37){
                        numDataPoints = Integer.parseInt(lineEntry);
                    } else if (q>37 && q<=37+numDataPoints){
                        passPattern.add(Double.parseDouble(lineEntry));
                    } else {
                        numPasses = Integer.parseInt(lineEntry);
                    }
                    q++;
                }
            } catch (Exception e){
                e.printStackTrace();
            }
            textFieldGS2.setText(line.get(29));
            textFieldWV2.setText(line.get(30));
            textFieldWD2.setText(line.get(31));
            pHA = 0;
            if(Double.parseDouble(line.get(32)) > 0) {
                pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
            } else {
                pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
            }
            textFieldPH2.setText(Double.toString(pHA));
            textFieldSH2.setText(line.get(34));

            patternObject.setPass2(passPattern);
            cBPass2.setSelected(true);
            isScannedPass2 = true;
            toggleButtonPass2.setDisable(false);
            toggleButtonPass2.setSelected(true);
        }

        //Read in Pass 3
        if(numPasses>2){
            line = new ArrayList<>();
            passPattern = new ArrayList<>();
            try {
                String datafileC = (datafile.toString().substring(0, datafile.toString().length() -1) + "C");
                BufferedReader in = new BufferedReader(new FileReader(datafileC));

                String lineEntry;
                int q = 0;
                while ((lineEntry = in.readLine()) != null) {
                    lineEntry = lineEntry.replace("\"",""); //Remove quotes
                    if(q<37){
                        line.add(lineEntry);
                    } else if (q==37){
                        numDataPoints = Integer.parseInt(lineEntry);
                    } else if (q>37 && q<=37+numDataPoints){
                        passPattern.add(Double.parseDouble(lineEntry));
                    } else {
                        numPasses = Integer.parseInt(lineEntry);
                    }
                    q++;
                }
            } catch (Exception e){
                e.printStackTrace();
            }
            textFieldGS3.setText(line.get(29));
            textFieldWV3.setText(line.get(30));
            textFieldWD3.setText(line.get(31));
            pHA = 0;
            if(Double.parseDouble(line.get(32)) > 0) {
                pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
            } else {
                pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
            }
            textFieldPH3.setText(Double.toString(pHA));
            textFieldSH3.setText(line.get(34));

            patternObject.setPass3(passPattern);
            cBPass3.setSelected(true);
            isScannedPass3 = true;
            toggleButtonPass3.setDisable(false);
            toggleButtonPass3.setSelected(true);
        }

        //Read in Pass 4
        if(numPasses>3){
            line = new ArrayList<>();
            passPattern = new ArrayList<>();
            try {
                String datafileD = (datafile.toString().substring(0, datafile.toString().length() -1) + "D");
                BufferedReader in = new BufferedReader(new FileReader(datafileD));

                String lineEntry;
                int q = 0;
                while ((lineEntry = in.readLine()) != null) {
                    lineEntry = lineEntry.replace("\"",""); //Remove quotes
                    if(q<37){
                        line.add(lineEntry);
                    } else if (q==37){
                        numDataPoints = Integer.parseInt(lineEntry);
                    } else if (q>37 && q<=37+numDataPoints){
                        passPattern.add(Double.parseDouble(lineEntry));
                    } else {
                        numPasses = Integer.parseInt(lineEntry);
                    }
                    q++;
                }
            } catch (Exception e){
                e.printStackTrace();
            }
            textFieldGS4.setText(line.get(29));
            textFieldWV4.setText(line.get(30));
            textFieldWD4.setText(line.get(31));
            pHA = 0;
            if(Double.parseDouble(line.get(32)) > 0) {
                pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
            } else {
                pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
            }
            textFieldPH4.setText(Double.toString(pHA));
            textFieldSH4.setText(line.get(34));

            patternObject.setPass4(passPattern);
            cBPass4.setSelected(true);
            isScannedPass4 = true;
            toggleButtonPass4.setDisable(false);
            toggleButtonPass4.setSelected(true);
        }

        //Read in Pass 5
        if(numPasses>4){
            line = new ArrayList<>();
            passPattern = new ArrayList<>();
            try {
                String datafileE = (datafile.toString().substring(0, datafile.toString().length() -1) + "E");
                BufferedReader in = new BufferedReader(new FileReader(datafileE));

                String lineEntry;
                int q = 0;
                while ((lineEntry = in.readLine()) != null) {
                    lineEntry = lineEntry.replace("\"",""); //Remove quotes
                    if(q<37){
                        line.add(lineEntry);
                    } else if (q==37){
                        numDataPoints = Integer.parseInt(lineEntry);
                    } else if (q>37 && q<=37+numDataPoints){
                        passPattern.add(Double.parseDouble(lineEntry));
                    } else {
                        numPasses = Integer.parseInt(lineEntry);
                    }
                    q++;
                }
            } catch (Exception e){
                e.printStackTrace();
            }
            textFieldGS5.setText(line.get(29));
            textFieldWV5.setText(line.get(30));
            textFieldWD5.setText(line.get(31));
            pHA = 0;
            if(Double.parseDouble(line.get(32)) > 0) {
                pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
            } else {
                pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
            }
            textFieldPH5.setText(Double.toString(pHA));
            textFieldSH5.setText(line.get(34));

            patternObject.setPass5(passPattern);
            cBPass5.setSelected(true);
            isScannedPass5 = true;
            toggleButtonPass5.setDisable(false);
            toggleButtonPass5.setSelected(true);
        }

        //Read in Pass 6
        if(numPasses>5){
            line = new ArrayList<>();
            passPattern = new ArrayList<>();
            try {
                String datafileF = (datafile.toString().substring(0, datafile.toString().length() -1) + "F");
                BufferedReader in = new BufferedReader(new FileReader(datafileF));

                String lineEntry;
                int q = 0;
                while ((lineEntry = in.readLine()) != null) {
                    lineEntry = lineEntry.replace("\"",""); //Remove quotes
                    if(q<37){
                        line.add(lineEntry);
                    } else if (q==37){
                        numDataPoints = Integer.parseInt(lineEntry);
                    } else if (q>37 && q<=37+numDataPoints){
                        passPattern.add(Double.parseDouble(lineEntry));
                    } else {
                        numPasses = Integer.parseInt(lineEntry);
                    }
                    q++;
                }
            } catch (Exception e){
                e.printStackTrace();
            }
            textFieldGS6.setText(line.get(29));
            textFieldWV6.setText(line.get(30));
            textFieldWD6.setText(line.get(31));
            pHA = 0;
            if(Double.parseDouble(line.get(32)) > 0) {
                pHA = (180 / (Math.PI)) * (Math.asin((Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) + (Double.parseDouble(line.get(31)));
            } else {
                pHA = Math.abs((180 / (Math.PI)) * (Math.asin(Math.abs(Double.parseDouble(line.get(32))) / (Double.parseDouble(line.get(30))))) - (Double.parseDouble(line.get(31))));
            }
            textFieldPH6.setText(Double.toString(pHA));
            textFieldSH6.setText(line.get(34));

            patternObject.setPass6(passPattern);
            cBPass6.setSelected(true);
            isScannedPass6 = true;
            toggleButtonPass6.setDisable(false);
            toggleButtonPass6.setSelected(true);
        }
        updateActivePasses();
        if (isScannedPass1 || isScannedPass2 || isScannedPass3 || isScannedPass4 || isScannedPass5 || isScannedPass6) {
            updatePass();
        }
        refreshPlots1();
        refreshPlots2();
        updatePassColors();

        labelCurrentFile.setText(datafile.toString());
    }

    private void openUSDA(File datafile){
        scrollPaneMain.setDisable(false);
        tabPaneMain.setDisable(false);
        clearAircraftAndSeriesData();
        clearPatternData();
        activePasses.clear();
        //Set Smoothing On
        toggleButtonSmoothSeries.setSelected(true);
        toggleButtonSmoothSeries.setText("Smooth Series: On");
        isSmoothSeries = true;
        toggleButtonSmoothAverage.setSelected(true);
        toggleButtonSmoothAverage.setText("Smooth Average: On");
        isSmoothAverage = true;

        labelSwathFinal.setText("65");
        setFlightlineLength(150);
        setSampleLength(0.10);
        resetHorizontalAxes(FLIGHTLINE_LENGTH);

        //Make a new pattern object with a sample length of 0.10 for uniformity
        patternObject = new PatternObject();
        patternObject.setSampleLength(0.10);
        //Split the file name ([Nxxxx] [Series Letter] [Pass Number])
        String[] parts = datafile.getName().split(" ");
        String reg = parts[0];
        String series = parts[1];
        String pass = parts[2];
        //Get all passes for this series
        File dir = new File(datafile.getParent());
        File[] files = dir.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name)
            {
                return name.startsWith(reg+" "+series) && name.endsWith(".txt");
            }
        });
        int numOfPasses = files.length;
        //Set Reg Num
        patternObject.setRegNumber(reg);
        textFieldRegNum.setText(reg);
        //Set Series Num
        String alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        for (int i = 0; i < parts[1].length(); i++) {
            patternObject.setSeriesNumber(alpha.indexOf(parts[1].charAt(i)) + 1);
            textFieldSeriesNum.setText(Integer.toString(patternObject.getSeriesNumber()));
        }
        //Check for PRN File
        File prn = new File(dir.getAbsolutePath()+"/Pilot Paramters.prn");
        String misc = "";
        if(prn.exists()){
            System.out.println(prn.getAbsolutePath()+" Exists");
            try {
                BufferedReader in = new BufferedReader(new FileReader(prn));

                String line;
                String[] entry;
                while ((line = in.readLine()) != null) {
                    //System.out.println(line);
                    entry = line.split("\t");
                    //System.out.println("..."+entry[1]+"...");
                    if (entry[0] != null && !entry[0].isEmpty() && entry[0].contains(reg)) {
                        if(entry.length>1 && entry[1]!=null){
                            textFieldPilot.setText(entry[1]);
                        }
                        if(entry.length>2 && entry[2]!=null){
                            textFieldStreet.setText(entry[2]);
                        }
                        if(entry.length>3 && entry[3]!=null){
                            textFieldCity.setText(entry[3]);
                        }
                        if(entry.length>4 && entry[4]!=null){
                            textFieldState.setText(entry[4]);
                        }
                        if(entry.length>5 && entry[5]!=null){
                            textFieldZIP.setText(entry[5]);
                        }
                        if(entry.length>7 && entry[7]!=null){
                            textFieldPhone.setText(entry[7]);
                        }
                        //Catch misc stuff
                        if(entry.length>6 && entry[6]!=null){
                            misc = misc + " " + entry[6];
                        }
                        if(entry.length>8 && entry[8]!=null){
                            misc = misc + " " + entry[8];
                        }
                        if(entry.length>9 && entry[9]!=null){
                            misc = misc + " " + entry[9];
                        }
                        textAreaNotes.setText(misc);
                    }
                }
                in.close();

            } catch (Exception e){
                e.printStackTrace();
            }
        }
        //Read In Pass 1
        if (numOfPasses>0) {
            patternObject.pass1 = DoubleStream.of(getUSDAPattern(files[0])).boxed().collect(Collectors.toCollection(ArrayList::new));
            patternObject.pass1Mod = trimAndZeroPattern(patternObject.pass1,0,0);
            isScannedPass1 = true;
            cBPass1.setSelected(true);
            toggleButtonPass1.setDisable(false);
            toggleButtonPass1.setSelected(true);
        }
        //Read In Pass 2
        if (numOfPasses>1) {
            patternObject.pass2 = DoubleStream.of(getUSDAPattern(files[1])).boxed().collect(Collectors.toCollection(ArrayList::new));
            patternObject.pass2Mod = trimAndZeroPattern(patternObject.pass2,0,0);
            isScannedPass2 = true;
            cBPass2.setSelected(true);
            toggleButtonPass2.setDisable(false);
            toggleButtonPass2.setSelected(true);
        }
        //Read In Pass 3
        if (numOfPasses>2) {
            patternObject.pass3 = DoubleStream.of(getUSDAPattern(files[2])).boxed().collect(Collectors.toCollection(ArrayList::new));
            patternObject.pass3Mod = trimAndZeroPattern(patternObject.pass3,0,0);
            isScannedPass3 = true;
            cBPass3.setSelected(true);
            toggleButtonPass3.setDisable(false);
            toggleButtonPass3.setSelected(true);
        }


        updateActivePasses();

        //plotOverlay();
        refreshPlots1();
        refreshPlots2();
    }

    private void openAccuStain(XSSFWorkbook wb){
        clearCardSlate();
        cardObject = new CardObject(DEFAULT_THRESHOLD);
        cardObject.getRm().setVisible(false);
        //Image Extraction
        List lst = wb.getAllPictures();
        for (Iterator it = lst.iterator(); it.hasNext(); ) {
            PictureData pict = (PictureData) it.next();
            byte[] data = pict.getData();
            try {
                BufferedImage bi = ImageIO.read(new ByteArrayInputStream(data));
                cardObject.getCardBufferedImages().add(bi);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //Get Card Spacing and Spread Factors
        XSSFSheet sh = wb.getSheet("Card Data");
        CARD_SPACING = (int) sh.getRow(0).getCell(1).getNumericCellValue();
        spinnerCardPassNum.getValueFactory().setValue((int) sh.getRow(1).getCell(1).getNumericCellValue());
        if(sh.getRow(2).getCell(1).getStringCellValue().contains("Direct")){
            setSpreadFactor(0);
        } else {
            setSpreadFactor(1);
        }
        setSpreadABC(sh.getRow(3).getCell(1).getNumericCellValue(),sh.getRow(4).getCell(1).getNumericCellValue(),sh.getRow(5).getCell(1).getNumericCellValue());
        //Get scanned cards and used cards
        for(int i=4; i<=12; i++){
            cardObject.getCardList().add(sh.getRow(1).getCell(i).getBooleanCellValue());
            cardObject.getUsingCardList().add(sh.getRow(2).getCell(i).getBooleanCellValue());
        }
        //Iterate through cards with updated thresholds
        int indexOfROI = 0;
        for(int i=0; i<cardObject.getCardList().size(); i++){
            if(cardObject.getCardList().get(i)){
                //cardObject.getCardImages().add(originalSnapshot(indexOfROI));
                cardObject.getCardNameList().add(sh.getRow(0).getCell(4+i).getStringCellValue());
                cardObject.getThresholds().set(indexOfROI,(int) sh.getRow(3).getCell(i+4).getNumericCellValue());
                //cardObject.getCardAreas().add(sh.getRow(6).getCell(i+4).getNumericCellValue());
                //cardObject.getCardImagesThresholded().add(thresholdedSnapshot(indexOfROI, cardObject.getThresholds().get(indexOfROI)));
                //logStains(getStainsFromCard(thresholdImage(indexOfROI,cardObject.getThresholds().get(indexOfROI))),indexOfROI);
                //displayCardOnScreen(indexOfROI,i,cardObject.getCardImages().get(indexOfROI),cardObject.getCardImagesThresholded().get(indexOfROI),cardObject.getThresholds().get(indexOfROI));

                indexOfROI++;
            }
        }
        iterateThroughCards(true);
        showComposites();

        clickCardToggle();

    }

    private void openAccuStainLegacy(XSSFWorkbook wb){
        XSSFSheet sh = wb.getSheet("Card Index");
        clearCardSlate();
        cardObject = new CardObject(DEFAULT_THRESHOLD);
        cardObject.getRm().setVisible(false);
        //Image Extraction
        List lst = wb.getAllPictures();
        for (Iterator it = lst.iterator(); it.hasNext(); ) {
            PictureData pict = (PictureData) it.next();
            byte[] data = pict.getData();
            try {
                BufferedImage bi = ImageIO.read(new ByteArrayInputStream(data));
                cardObject.getCardBufferedImages().add(bi);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //Get Card Spacing and Spread Factors
        CARD_SPACING = (int) sh.getRow(0).getCell(1).getNumericCellValue();
        if (sh.getRow(1).getCell(1).getCellType()==Cell.CELL_TYPE_NUMERIC) {
            spinnerCardPassNum.getValueFactory().setValue((int) sh.getRow(1).getCell(1).getNumericCellValue());
        } else {
            spinnerCardPassNum.getValueFactory().setValue(Integer.parseInt(sh.getRow(1).getCell(1).getStringCellValue()));
        }
        for(int i=0; i<9; i++){
            cardObject.getCardList().add(sh.getRow(5+i).getCell(0).getBooleanCellValue());
            cardObject.getUsingCardList().add(sh.getRow(5+i).getCell(0).getBooleanCellValue());
            if(wb.getSheetAt(5+i)!=null){
                cardObject.getCardNameList().add(wb.getSheetAt(5+i).getSheetName());
            }
        }
        iterateThroughCards(false);
        showComposites();
        clickCardToggle();
        //Update to new format

        while(wb.getNumberOfSheets()>4){
            wb.removeSheetAt(wb.getNumberOfSheets()-1);
        }
        try {
            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        saveAllCardDataToDataFile();
    }

    public double[] getUSDAPattern(File passfile){
        //double[] emmsAI = new double[1501];
        try {
            BufferedReader in = new BufferedReader(new FileReader(passfile));
            ArrayList<Double> locs = new ArrayList<>();
            ArrayList<Double> emms = new ArrayList<>();
            String line;
            String[] entry;
            int q = 0;
            while ((line = in.readLine()) != null) {
                //System.out.println(line);
                entry = line.split("\\s+");
                //System.out.println("..."+entry[1]+"...");
                if (entry.length>3 && q>1 && entry[1] != null) {

                    locs.add(Double.parseDouble(entry[1]));
                    emms.add(Double.parseDouble(entry[2]));
                }
                q++;
            }
            in.close();
            //Convert locs and emms to arrays for future interpolation
            double[] locsA = new double[locs.size()];
            for (int i = 0; i < locsA.length; i++) {
                locsA[i] = locs.get(i);
            }
            double[] emmsA = new double[emms.size()];
            for (int i = 0; i < emmsA.length; i++) {
                emmsA[i] = emms.get(i);
            }

            LinearInterpolator li = new LinearInterpolator();
            PolynomialSplineFunction psf = li.interpolate(locsA, emmsA);
            //Interpolate data onto a 0.10 ft scale over 150 ft
            double fl = locs.get(locs.size()-1);
            double sl = 0.10; //sample length to interpolate to
            double[] locsAI = new double[(int) Math.round(fl/sl)];
            double[] emmsAI = new double[(int) Math.round(fl/sl)];
            System.out.println("fl = "+fl+" locsAI.size = "+locsAI.length);
            for (int i = 0; i < locsAI.length; i++) {
                locsAI[i] = i * 0.10;
                emmsAI[i] = psf.value(i*0.10);
            }
            return emmsAI;
        } catch (Exception e){
            e.printStackTrace();
            return null;
        }

    }

    public void updateDataFile(){

        //Append Trims
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("Pattern Data");
            sh.getRow(0).createCell(2).setCellValue(patternObject.getTrim1L());
            sh.getRow(0).createCell(4).setCellValue(patternObject.getTrim2L());
            sh.getRow(0).createCell(6).setCellValue(patternObject.getTrim3L());
            sh.getRow(0).createCell(8).setCellValue(patternObject.getTrim4L());
            sh.getRow(0).createCell(10).setCellValue(patternObject.getTrim5L());
            sh.getRow(0).createCell(12).setCellValue(patternObject.getTrim6L());
            sh.getRow(1).createCell(2).setCellValue(patternObject.getTrim1R());
            sh.getRow(1).createCell(4).setCellValue(patternObject.getTrim2R());
            sh.getRow(1).createCell(6).setCellValue(patternObject.getTrim3R());
            sh.getRow(1).createCell(8).setCellValue(patternObject.getTrim4R());
            sh.getRow(1).createCell(10).setCellValue(patternObject.getTrim5R());
            sh.getRow(1).createCell(12).setCellValue(patternObject.getTrim6R());
            sh.getRow(2).createCell(2).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(4).setCellValue(patternObject.getTrimVertical2());
            sh.getRow(2).createCell(6).setCellValue(patternObject.getTrimVertical3());
            sh.getRow(2).createCell(8).setCellValue(patternObject.getTrimVertical4());
            sh.getRow(2).createCell(10).setCellValue(patternObject.getTrimVertical5());
            sh.getRow(2).createCell(12).setCellValue(patternObject.getTrimVertical6());
            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}

        try {
            //Create XLSX of Data
            FileInputStream fis = new FileInputStream(currentDataFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);

            //Create Third Sheet for Report Headers
            XSSFSheet sh = wb.getSheet("Fly-In Data");

            if (header1 != null && !header1.isEmpty()) {
                sh.createRow(0).createCell(0).setCellValue(header1);
            } else {
                sh.createRow(0).createCell(0);
            }
            if (header2 != null && !header2.isEmpty()) {
                sh.createRow(1).createCell(0).setCellValue(header2);
            } else {
                sh.createRow(1).createCell(0);
            }
            if (header3 != null && !header3.isEmpty()) {
                sh.createRow(2).createCell(0).setCellValue(header3);
            } else {
                sh.createRow(2).createCell(0);
            }
            if (header4 != null && !header4.isEmpty()) {
                sh.createRow(3).createCell(0).setCellValue(header4);
            } else {
                sh.createRow(3).createCell(0);
            }

            sh.autoSizeColumn(0);

            //Create First Sheet for Parameters
            XSSFSheet sh0 = wb.getSheet("Aircraft Data");

            XSSFRow row0 = sh0.createRow(0);
            row0.createCell(0).setCellValue("Pilot:");
            row0.createCell(1).setCellValue(textFieldPilot.getText());
            row0.createCell(2).setCellValue(textFieldEmail.getText());

            XSSFRow row1 = sh0.createRow(1);
            row1.createCell(0).setCellValue("Business Name:");
            row1.createCell(1).setCellValue(textFieldBusinessName.getText());

            XSSFRow row2 = sh0.createRow(2);
            row2.createCell(0).setCellValue("Phone:");
            row2.createCell(1).setCellValue(textFieldPhone.getText());

            XSSFRow row3 = sh0.createRow(3);
            row3.createCell(0).setCellValue("Street:");
            row3.createCell(1).setCellValue(textFieldStreet.getText());

            XSSFRow row4 = sh0.createRow(4);
            row4.createCell(0).setCellValue("Ciy:");
            row4.createCell(1).setCellValue(textFieldCity.getText());

            XSSFRow row5 = sh0.createRow(5);
            row5.createCell(0).setCellValue("State:");
            row5.createCell(1).setCellValue(textFieldState.getText());
            row5.createCell(2).setCellValue(textFieldZIP.getText());

            XSSFRow row6 = sh0.createRow(6);
            row6.createCell(0).setCellValue("Reg. #:");
            row6.createCell(1).setCellValue(textFieldRegNum.getText());

            XSSFRow row7 = sh0.createRow(7);
            row7.createCell(0).setCellValue("Series #::");
            row7.createCell(1).setCellValue(textFieldSeriesNum.getText());

            XSSFRow row8 = sh0.createRow(8);
            row8.createCell(0).setCellValue("Aircraft Make:");
            row8.createCell(1).setCellValue(comboBoxAircraftMake.getValue().toString());

            XSSFRow row9 = sh0.createRow(9);
            row9.createCell(0).setCellValue("Aircraft Model:");
            row9.createCell(1).setCellValue(comboBoxAircraftModel.getValue().toString());

            XSSFRow row10 = sh0.createRow(10);
            row10.createCell(0).setCellValue("#1 Nozzle Type:");
            if(!comboBoxNozzle1Type.getValue().equals("Select Nozzle")) {
                row10.createCell(1).setCellValue(comboBoxNozzle1Type.getValue().toString());
            }

            XSSFRow row11 = sh0.createRow(11);
            row11.createCell(0).setCellValue("#1 Nozzle Quantity:");
            row11.createCell(1).setCellValue(textFieldNozzle1Quant.getText());

            XSSFRow row12 = sh0.createRow(12);
            row12.createCell(0).setCellValue("#1 Nozzle Size:");
            if(!comboBoxNozzle1Size.getValue().equals("N/A")) {
                row12.createCell(1).setCellValue(comboBoxNozzle1Size.getValue().toString());
            }

            XSSFRow row13 = sh0.createRow(13);
            row13.createCell(0).setCellValue("#1 Nozzle Def:");
            if(!comboBoxNozzle1Def.getValue().equals("N/A")) {
                row13.createCell(1).setCellValue(comboBoxNozzle1Def.getValue().toString());
            }

            XSSFRow row14 = sh0.createRow(14);
            row14.createCell(0).setCellValue("#2 Nozzle Type:");
            if(!comboBoxNozzle2Type.getValue().equals("Select Nozzle")) {
                row14.createCell(1).setCellValue(comboBoxNozzle2Type.getValue().toString());
            }

            XSSFRow row15 = sh0.createRow(15);
            row15.createCell(0).setCellValue("#2 Nozzle Quantity::");
            row15.createCell(1).setCellValue(textFieldNozzle2Quant.getText());

            XSSFRow row16 = sh0.createRow(16);
            row16.createCell(0).setCellValue("#2 Nozzle Size:");
            if(!comboBoxNozzle2Size.getValue().equals("N/A")) {
                row16.createCell(1).setCellValue(comboBoxNozzle2Size.getValue().toString());
            }

            XSSFRow row17 = sh0.createRow(17);
            row17.createCell(0).setCellValue("#2 Nozzle Def:");
            if(!comboBoxNozzle2Def.getValue().equals("N/A")) {
                row17.createCell(1).setCellValue(comboBoxNozzle2Def.getValue().toString());
            }

            XSSFRow row18 = sh0.createRow(18);
            if (!METRIC) {
                row18.createCell(0).setCellValue("Boom Pressure (PSI):");
            } else {
                row18.createCell(0).setCellValue("Boom Pressure (PSI):");
            }
            row18.createCell(1).setCellValue(textFieldBoomPressure.getText());

            XSSFRow row19 = sh0.createRow(19);
            if (!METRIC) {
                row19.createCell(0).setCellValue("Target Rate (GPA):");
            } else {
                row19.createCell(0).setCellValue("Target Rate (L/ha):");
            }
            row19.createCell(1).setCellValue(textFieldTargetRate.getText());

            XSSFRow row20 = sh0.createRow(20);
            if (!METRIC) {
                row20.createCell(0).setCellValue("Target Swath (FT):");
            } else {
                row20.createCell(0).setCellValue("Target Swath (m):");
            }
            row20.createCell(1).setCellValue(textFieldTargetSwath.getText());

            XSSFRow row26 = sh0.createRow(26);
            row26.createCell(0).setCellValue("Time:");
            row26.createCell(1).setCellValue(textFieldTime.getText());

            XSSFRow row27 = sh0.createRow(27);
            row27.createCell(0).setCellValue("Wing Span:");
            row27.createCell(1).setCellValue(textFieldWingSpan.getText());

            XSSFRow row28 = sh0.createRow(28);
            row28.createCell(0).setCellValue("Boom Width:");
            row28.createCell(1).setCellValue(textFieldBoomWidth.getText());

            XSSFRow row29 = sh0.createRow(29);
            row29.createCell(0).setCellValue("% Boom Width:");
            if (!textFieldWingSpan.getText().isEmpty() && !textFieldBoomWidth.getText().isEmpty()) {
                double ws = Double.parseDouble(textFieldWingSpan.getText());
                double bw = Double.parseDouble(textFieldBoomWidth.getText());
                double percentBW = (bw/ws)*100;
                if(percentBW >=0 && percentBW <=100) {
                    row29.createCell(1).setCellValue(Double.toString(percentBW));
                }
            }

            XSSFRow row30 = sh0.createRow(30);
            if (!METRIC) {
                row30.createCell(0).setCellValue("Boom Drop (IN):");
            } else {
                row30.createCell(0).setCellValue("Boom Drop (cm):");
            }
            row30.createCell(1).setCellValue(textFieldBoomDrop.getText());

            XSSFRow row31 = sh0.createRow(31);
            if (!METRIC) {
                row31.createCell(0).setCellValue("Nozzle Spacing (IN):");
            } else {
                row31.createCell(0).setCellValue("Nozzle Spacing (cm):");
            }
            row31.createCell(1).setCellValue(textFieldNozzleSpacing.getText());

            XSSFRow row32 = sh0.createRow(32);
            row32.createCell(0).setCellValue("Winglets?");
            row32.createCell(1).setCellValue(comboBoxWinglets.getSelectionModel().getSelectedItem().toString());

            XSSFRow row33 = sh0.createRow(33);
            row33.createCell(0).setCellValue("Notes:");
            row33.createCell(1).setCellValue(textAreaNotes.getText());


            XSSFRow row35 = sh0.createRow(35);
            row35.createCell(0).setCellValue("METRIC");
            row35.createCell(1).setCellValue(METRIC);

            sh0.autoSizeColumn(0);
            sh0.autoSizeColumn(1);

            //Create Tab for Series Data
            XSSFSheet sh1 = wb.getSheet("Series Data");

            XSSFRow sh1row0 = sh1.createRow(0);
            sh1row0.createCell(1).setCellValue("Pass 1");
            sh1row0.createCell(2).setCellValue("Pass 2");
            sh1row0.createCell(3).setCellValue("Pass 3");
            sh1row0.createCell(4).setCellValue("Pass 4");
            sh1row0.createCell(5).setCellValue("Pass 5");
            sh1row0.createCell(6).setCellValue("Pass 6");

            XSSFRow sh1row1 = sh1.createRow(1);
            sh1row1.createCell(0).setCellValue("Ground Speed (MPH)");
            sh1row1.createCell(1).setCellValue(textFieldGS1.getText());
            sh1row1.createCell(2).setCellValue(textFieldGS2.getText());
            sh1row1.createCell(3).setCellValue(textFieldGS3.getText());
            sh1row1.createCell(4).setCellValue(textFieldGS4.getText());
            sh1row1.createCell(5).setCellValue(textFieldGS5.getText());
            sh1row1.createCell(6).setCellValue(textFieldGS6.getText());

            XSSFRow sh1row2 = sh1.createRow(2);
            if (!METRIC) {
                sh1row2.createCell(0).setCellValue("Spray Height (FT)");
            } else {
                sh1row2.createCell(0).setCellValue("Spray Height (m)");
            }
            sh1row2.createCell(1).setCellValue(textFieldSH1.getText());
            sh1row2.createCell(2).setCellValue(textFieldSH2.getText());
            sh1row2.createCell(3).setCellValue(textFieldSH3.getText());
            sh1row2.createCell(4).setCellValue(textFieldSH4.getText());
            sh1row2.createCell(5).setCellValue(textFieldSH5.getText());
            sh1row2.createCell(6).setCellValue(textFieldSH6.getText());

            XSSFRow sh1row3 = sh1.createRow(3);
            if (invertPH) {
                sh1row3.createCell(0).setCellValue("Pass Heading (Deg)*");
            } else {
                sh1row3.createCell(0).setCellValue("Pass Heading (Deg)");
            }
            sh1row3.createCell(1).setCellValue(textFieldPH1.getText());
            sh1row3.createCell(2).setCellValue(textFieldPH2.getText());
            sh1row3.createCell(3).setCellValue(textFieldPH3.getText());
            sh1row3.createCell(4).setCellValue(textFieldPH4.getText());
            sh1row3.createCell(5).setCellValue(textFieldPH5.getText());
            sh1row3.createCell(6).setCellValue(textFieldPH6.getText());

            XSSFRow sh1row4 = sh1.createRow(4);
            sh1row4.createCell(0).setCellValue("Wind Direction (Deg)");
            sh1row4.createCell(1).setCellValue(textFieldWD1.getText());
            sh1row4.createCell(2).setCellValue(textFieldWD2.getText());
            sh1row4.createCell(3).setCellValue(textFieldWD3.getText());
            sh1row4.createCell(4).setCellValue(textFieldWD4.getText());
            sh1row4.createCell(5).setCellValue(textFieldWD5.getText());
            sh1row4.createCell(6).setCellValue(textFieldWD6.getText());

            XSSFRow sh1row5 = sh1.createRow(5);
            sh1row5.createCell(0).setCellValue("Wind Velocity (MPH)");
            sh1row5.createCell(1).setCellValue(textFieldWV1.getText());
            sh1row5.createCell(2).setCellValue(textFieldWV2.getText());
            sh1row5.createCell(3).setCellValue(textFieldWV3.getText());
            sh1row5.createCell(4).setCellValue(textFieldWV4.getText());
            sh1row5.createCell(5).setCellValue(textFieldWV5.getText());
            sh1row5.createCell(6).setCellValue(textFieldWV6.getText());

            XSSFRow sh1row6 = sh1.createRow(6);
            if (!METRIC) {
                sh1row6.createCell(0).setCellValue("Ambient Temperature (Deg F)");
            } else {
                sh1row6.createCell(0).setCellValue("Ambient Temperature (Deg C)");
            }
            sh1row6.createCell(1).setCellValue(textFieldAT.getText());

            XSSFRow sh1row7 = sh1.createRow(7);
            sh1row7.createCell(0).setCellValue("Relative Humidity (%)");
            sh1row7.createCell(1).setCellValue(textFieldRH.getText());

            sh1.autoSizeColumn(0);
            sh1.autoSizeColumn(1);
            sh1.autoSizeColumn(2);
            sh1.autoSizeColumn(3);
            sh1.autoSizeColumn(4);
            sh1.autoSizeColumn(5);
            sh1.autoSizeColumn(6);

            try {
                wb.write(new FileOutputStream(currentDataFile));
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("File Save Success");
                alert.setHeaderText("Successfully updated data file");
                alert.setContentText("File Location:\n"+currentDataFile);
                alert.showAndWait();
            }  catch (Exception e){
                e.printStackTrace();
                Alert alert0 = new Alert(Alert.AlertType.ERROR);
                alert0.setTitle("File Save Error");
                alert0.setHeaderText("Failed to update file:\n"+currentDataFile);
                alert0.setContentText("If file is open in another program, exit that program and try again.");
                alert0.showAndWait();
            }
        } catch (IOException e) {
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("File Open Error");
            alert.setHeaderText("Failed to open file:\n"+currentDataFile);
            alert.setContentText("Check File contents and try again.");
            alert.showAndWait();
        }
    }                            //XLSX

    public void saveCurrentAircraftAndSeriesData() {
        //Create XLSX of Data
        XSSFWorkbook wb = new XSSFWorkbook();

        //Create Sheet for Report Headers
        XSSFSheet sh = wb.createSheet("Fly-In Data");

        if (header1 != null && !header1.isEmpty()) {
            sh.createRow(0).createCell(0).setCellValue(header1);
        } else {
            sh.createRow(0).createCell(0);
        }
        if (header2 != null && !header2.isEmpty()) {
            sh.createRow(1).createCell(0).setCellValue(header2);
        } else {
            sh.createRow(1).createCell(0);
        }
        if (header3 != null && !header3.isEmpty()) {
            sh.createRow(2).createCell(0).setCellValue(header3);
        } else {
            sh.createRow(2).createCell(0);
        }
        if (header4 != null && !header4.isEmpty()) {
            sh.createRow(3).createCell(0).setCellValue(header4);
        } else {
            sh.createRow(3).createCell(0);
        }

        sh.autoSizeColumn(0);

        //Create First Sheet for Parameters
        XSSFSheet sh0 = wb.createSheet("Aircraft Data");

        XSSFRow row0 = sh0.createRow(0);
        row0.createCell(0).setCellValue("Pilot:");
        row0.createCell(1).setCellValue(textFieldPilot.getText());
        row0.createCell(2).setCellValue(textFieldEmail.getText());

        XSSFRow row1 = sh0.createRow(1);
        row1.createCell(0).setCellValue("Business Name:");
        row1.createCell(1).setCellValue(textFieldBusinessName.getText());

        XSSFRow row2 = sh0.createRow(2);
        row2.createCell(0).setCellValue("Phone:");
        row2.createCell(1).setCellValue(textFieldPhone.getText());

        XSSFRow row3 = sh0.createRow(3);
        row3.createCell(0).setCellValue("Street:");
        row3.createCell(1).setCellValue(textFieldStreet.getText());

        XSSFRow row4 = sh0.createRow(4);
        row4.createCell(0).setCellValue("Ciy:");
        row4.createCell(1).setCellValue(textFieldCity.getText());

        XSSFRow row5 = sh0.createRow(5);
        row5.createCell(0).setCellValue("State:");
        row5.createCell(1).setCellValue(textFieldState.getText());
        row5.createCell(2).setCellValue(textFieldZIP.getText());

        XSSFRow row6 = sh0.createRow(6);
        row6.createCell(0).setCellValue("Reg. #:");
        row6.createCell(1).setCellValue(textFieldRegNum.getText());

        XSSFRow row7 = sh0.createRow(7);
        row7.createCell(0).setCellValue("Series #::");
        row7.createCell(1).setCellValue(textFieldSeriesNum.getText());

        XSSFRow row8 = sh0.createRow(8);
        row8.createCell(0).setCellValue("Aircraft Make:");
        row8.createCell(1).setCellValue(comboBoxAircraftMake.getValue().toString());

        XSSFRow row9 = sh0.createRow(9);
        row9.createCell(0).setCellValue("Aircraft Model:");
        row9.createCell(1).setCellValue(comboBoxAircraftModel.getValue().toString());

        XSSFRow row10 = sh0.createRow(10);
        row10.createCell(0).setCellValue("#1 Nozzle Type:");
        if(!comboBoxNozzle1Type.getValue().equals("Select Nozzle")) {
            row10.createCell(1).setCellValue(comboBoxNozzle1Type.getValue().toString());
        }

        XSSFRow row11 = sh0.createRow(11);
        row11.createCell(0).setCellValue("#1 Nozzle Quantity:");
        row11.createCell(1).setCellValue(textFieldNozzle1Quant.getText());

        XSSFRow row12 = sh0.createRow(12);
        row12.createCell(0).setCellValue("#1 Nozzle Size:");
        if(!comboBoxNozzle1Size.getValue().equals("N/A")) {
            row12.createCell(1).setCellValue(comboBoxNozzle1Size.getValue().toString());
        }

        XSSFRow row13 = sh0.createRow(13);
        row13.createCell(0).setCellValue("#1 Nozzle Def:");
        if(!comboBoxNozzle1Def.getValue().equals("N/A")) {
            row13.createCell(1).setCellValue(comboBoxNozzle1Def.getValue().toString());
        }

        XSSFRow row14 = sh0.createRow(14);
        row14.createCell(0).setCellValue("#2 Nozzle Type:");
        if(!comboBoxNozzle2Type.getValue().equals("Select Nozzle")) {
            row14.createCell(1).setCellValue(comboBoxNozzle2Type.getValue().toString());
        }

        XSSFRow row15 = sh0.createRow(15);
        row15.createCell(0).setCellValue("#2 Nozzle Quantity::");
        row15.createCell(1).setCellValue(textFieldNozzle2Quant.getText());

        XSSFRow row16 = sh0.createRow(16);
        row16.createCell(0).setCellValue("#2 Nozzle Size:");
        if(!comboBoxNozzle2Size.getValue().equals("N/A")) {
            row16.createCell(1).setCellValue(comboBoxNozzle2Size.getValue().toString());
        }

        XSSFRow row17 = sh0.createRow(17);
        row17.createCell(0).setCellValue("#2 Nozzle Def:");
        if(!comboBoxNozzle2Def.getValue().equals("N/A")) {
            row17.createCell(1).setCellValue(comboBoxNozzle2Def.getValue().toString());
        }

        XSSFRow row18 = sh0.createRow(18);
        if (!METRIC) {
            row18.createCell(0).setCellValue("Boom Pressure (PSI):");
        } else {
            row18.createCell(0).setCellValue("Boom Pressure (PSI):");
        }
        row18.createCell(1).setCellValue(textFieldBoomPressure.getText());

        XSSFRow row19 = sh0.createRow(19);
        if (!METRIC) {
            row19.createCell(0).setCellValue("Target Rate (GPA):");
        } else {
            row19.createCell(0).setCellValue("Target Rate (L/ha):");
        }
        row19.createCell(1).setCellValue(textFieldTargetRate.getText());

        XSSFRow row20 = sh0.createRow(20);
        if (!METRIC) {
            row20.createCell(0).setCellValue("Target Swath (FT):");
        } else {
            row20.createCell(0).setCellValue("Target Swath (m):");
        }
        row20.createCell(1).setCellValue(textFieldTargetSwath.getText());

        XSSFRow row26 = sh0.createRow(26);
        row26.createCell(0).setCellValue("Time:");
        row26.createCell(1).setCellValue(textFieldTime.getText());

        XSSFRow row27 = sh0.createRow(27);
        row27.createCell(0).setCellValue("Wing Span:");
        row27.createCell(1).setCellValue(textFieldWingSpan.getText());

        XSSFRow row28 = sh0.createRow(28);
        row28.createCell(0).setCellValue("Boom Width:");
        row28.createCell(1).setCellValue(textFieldBoomWidth.getText());

        XSSFRow row29 = sh0.createRow(29);
        row29.createCell(0).setCellValue("% Boom Width:");
        if (!textFieldWingSpan.getText().isEmpty() && !textFieldBoomWidth.getText().isEmpty()) {
            double ws = Double.parseDouble(textFieldWingSpan.getText());
            double bw = Double.parseDouble(textFieldBoomWidth.getText());
            double percentBW = (bw/ws)*100;
            if(percentBW >=0 && percentBW <=100) {
                row29.createCell(1).setCellValue(Double.toString(percentBW));
            }
        }

        XSSFRow row30 = sh0.createRow(30);
        if (!METRIC) {
            row30.createCell(0).setCellValue("Boom Drop (IN):");
        } else {
            row30.createCell(0).setCellValue("Boom Drop (cm):");
        }
        row30.createCell(1).setCellValue(textFieldBoomDrop.getText());

        XSSFRow row31 = sh0.createRow(31);
        if (!METRIC) {
            row31.createCell(0).setCellValue("Nozzle Spacing (IN):");
        } else {
            row31.createCell(0).setCellValue("Nozzle Spacing (cm):");
        }
        row31.createCell(1).setCellValue(textFieldNozzleSpacing.getText());

        XSSFRow row32 = sh0.createRow(32);
        row32.createCell(0).setCellValue("Winglets?");
        row32.createCell(1).setCellValue(comboBoxWinglets.getSelectionModel().getSelectedItem().toString());

        XSSFRow row33 = sh0.createRow(33);
        row33.createCell(0).setCellValue("Notes:");
        row33.createCell(1).setCellValue(textAreaNotes.getText());


        XSSFRow row35 = sh0.createRow(35);
        row35.createCell(0).setCellValue("METRIC");
        row35.createCell(1).setCellValue(METRIC);

        sh0.autoSizeColumn(0);
        sh0.autoSizeColumn(1);

        //Create Tab for Series Data
        XSSFSheet sh1 = wb.createSheet("Series Data");

        XSSFRow sh1row0 = sh1.createRow(0);
        sh1row0.createCell(1).setCellValue("Pass 1");
        sh1row0.createCell(2).setCellValue("Pass 2");
        sh1row0.createCell(3).setCellValue("Pass 3");
        sh1row0.createCell(4).setCellValue("Pass 4");
        sh1row0.createCell(5).setCellValue("Pass 5");
        sh1row0.createCell(6).setCellValue("Pass 6");

        XSSFRow sh1row1 = sh1.createRow(1);
        sh1row1.createCell(0).setCellValue("Ground Speed (MPH)");
        sh1row1.createCell(1).setCellValue(textFieldGS1.getText());
        sh1row1.createCell(2).setCellValue(textFieldGS2.getText());
        sh1row1.createCell(3).setCellValue(textFieldGS3.getText());
        sh1row1.createCell(4).setCellValue(textFieldGS4.getText());
        sh1row1.createCell(5).setCellValue(textFieldGS5.getText());
        sh1row1.createCell(6).setCellValue(textFieldGS6.getText());

        XSSFRow sh1row2 = sh1.createRow(2);
        if (!METRIC) {
            sh1row2.createCell(0).setCellValue("Spray Height (FT)");
        } else {
            sh1row2.createCell(0).setCellValue("Spray Height (m)");
        }
        sh1row2.createCell(1).setCellValue(textFieldSH1.getText());
        sh1row2.createCell(2).setCellValue(textFieldSH2.getText());
        sh1row2.createCell(3).setCellValue(textFieldSH3.getText());
        sh1row2.createCell(4).setCellValue(textFieldSH4.getText());
        sh1row2.createCell(5).setCellValue(textFieldSH5.getText());
        sh1row2.createCell(6).setCellValue(textFieldSH6.getText());

        XSSFRow sh1row3 = sh1.createRow(3);
        if (invertPH) {
            sh1row3.createCell(0).setCellValue("Pass Heading (Deg)*");
        } else {
            sh1row3.createCell(0).setCellValue("Pass Heading (Deg)");
        }
        sh1row3.createCell(1).setCellValue(textFieldPH1.getText());
        sh1row3.createCell(2).setCellValue(textFieldPH2.getText());
        sh1row3.createCell(3).setCellValue(textFieldPH3.getText());
        sh1row3.createCell(4).setCellValue(textFieldPH4.getText());
        sh1row3.createCell(5).setCellValue(textFieldPH5.getText());
        sh1row3.createCell(6).setCellValue(textFieldPH6.getText());

        XSSFRow sh1row4 = sh1.createRow(4);
        sh1row4.createCell(0).setCellValue("Wind Direction (Deg)");
        sh1row4.createCell(1).setCellValue(textFieldWD1.getText());
        sh1row4.createCell(2).setCellValue(textFieldWD2.getText());
        sh1row4.createCell(3).setCellValue(textFieldWD3.getText());
        sh1row4.createCell(4).setCellValue(textFieldWD4.getText());
        sh1row4.createCell(5).setCellValue(textFieldWD5.getText());
        sh1row4.createCell(6).setCellValue(textFieldWD6.getText());

        XSSFRow sh1row5 = sh1.createRow(5);
        sh1row5.createCell(0).setCellValue("Wind Velocity (MPH)");
        sh1row5.createCell(1).setCellValue(textFieldWV1.getText());
        sh1row5.createCell(2).setCellValue(textFieldWV2.getText());
        sh1row5.createCell(3).setCellValue(textFieldWV3.getText());
        sh1row5.createCell(4).setCellValue(textFieldWV4.getText());
        sh1row5.createCell(5).setCellValue(textFieldWV5.getText());
        sh1row5.createCell(6).setCellValue(textFieldWV6.getText());

        XSSFRow sh1row6 = sh1.createRow(6);
        if (!METRIC) {
            sh1row6.createCell(0).setCellValue("Ambient Temperature (Deg F)");
        } else {
            sh1row6.createCell(0).setCellValue("Ambient Temperature (Deg C)");
        }
        sh1row6.createCell(1).setCellValue(textFieldAT.getText());

        XSSFRow sh1row7 = sh1.createRow(7);
        sh1row7.createCell(0).setCellValue("Relative Humidity (%)");
        sh1row7.createCell(1).setCellValue(textFieldRH.getText());

        sh1.autoSizeColumn(0);
        sh1.autoSizeColumn(1);
        sh1.autoSizeColumn(2);
        sh1.autoSizeColumn(3);
        sh1.autoSizeColumn(4);
        sh1.autoSizeColumn(5);
        sh1.autoSizeColumn(6);

        //Create Tab for Pattern Data
        XSSFSheet sh2 = wb.createSheet("Pattern Data");
        sh2.createRow(0).createCell(1).setCellValue("Pass 1");
        sh2.getRow(0).createCell(3).setCellValue("Pass 2");
        sh2.getRow(0).createCell(5).setCellValue("Pass 3");
        sh2.getRow(0).createCell(7).setCellValue("Pass 4");
        sh2.getRow(0).createCell(9).setCellValue("Pass 5");
        sh2.getRow(0).createCell(11).setCellValue("Pass 6");
        sh2.createRow(1).createCell(0).setCellValue("String Length:");
        sh2.createRow(2).createCell(0).setCellValue("Sample Length:");
        sh2.createRow(3).createCell(0).setCellValue("Integration Time:");
        sh2.createRow(4).createCell(0).setCellValue("Ex/Em (nm(px)):");
        sh2.autoSizeColumn(0);
        for (int i = 0; i<=(FLIGHTLINE_LENGTH/SAMPLE_LENGTH); i++){
            sh2.createRow(i+5).createCell(0).setCellValue(-(FLIGHTLINE_LENGTH/2)+(SAMPLE_LENGTH*i));
        }

        try {
            FileChooser fc = new FileChooser();
            if (getCurrentDirectory()!=null) {
                if (getCurrentDirectory().exists()) {
                    fc.setInitialDirectory(getCurrentDirectory());
                }
            }
            if(textFieldSeriesNum.getLength() == 1) {
                fc.setInitialFileName(textFieldRegNum.getText() + " 0" + textFieldSeriesNum.getText()+".xlsx");
            } else if(textFieldSeriesNum.getLength() > 1){
                fc.setInitialFileName(textFieldRegNum.getText() + " " + textFieldSeriesNum.getText()+".xlsx");
            } else {
                fc.setInitialFileName(textFieldRegNum.getText()+".xlsx");
            }
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("AccuPatt", "*.xlsx");
            fc.getExtensionFilters().add(extFilter);

            if ((currentDataFile = fc.showSaveDialog(null)) != null) {
                wb.write(new FileOutputStream(currentDataFile));
                labelCurrentFile.setText(currentDataFile.toString());
                if (currentDataFile.getParentFile().exists() && currentDataFile.getParentFile().isDirectory()) {
                    setCurrentDirectory(currentDataFile.getParentFile());
                }
            } else {
                buttonRun.setDisable(false);
                scrollPaneMain.setDisable(false);
                tabPaneMain.setDisable(true);
            }

        }  catch (Exception e){
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("File Save Error");
            alert.setHeaderText("Unable to Write File");
            alert.setContentText("Check if file is open in another program and try again.");
            alert.showAndWait();
        }
        updateDirectoryInParams();
        generateFlyInReport();
    }         //XLSX

    public void savePrintedSwathWidth() {
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("Aircraft Data");
            if (sh.getRow(20).getCell(2)==null) {
                sh.getRow(20).createCell(2).setCellValue(labelSwathFinal.getText());
            } else {
                sh.getRow(20).getCell(2).setCellValue(labelSwathFinal.getText());
            }
            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}
    }                    //XLSX

    public void generateFlyInReport(){
        //Check if Report Exists. If not, create one. Else, open it and add to it.
        if(reportFile==null){
            File check = new File(currentDirectory+"/SAFEReport "+currentDirectory.getName()+".xlsx");
            if(check.exists()){
                reportFile=check;
            }
        }
        if(reportFile==null){
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = wb.createSheet("Fly-In Data");

            sh.createRow(0).createCell(1).setCellValue("Event:");
            sh.getRow(0).createCell(2).setCellValue(header1);

            sh.createRow(1).createCell(1).setCellValue("Location:");
            sh.getRow(1).createCell(2).setCellValue(header2);

            sh.createRow(2).createCell(1).setCellValue("Date:");
            sh.getRow(2).createCell(2).setCellValue(header3);

            sh.createRow(3).createCell(1).setCellValue("Analyst:");
            sh.getRow(3).createCell(2).setCellValue(header4);

            sh.createRow(4).createCell(1).setCellValue("# of Passes:");
            sh.getRow(4).createCell(2).setCellValue(0);

            XSSFRow row5 = sh.createRow(5);
            row5.createCell(0).setCellValue("Pilot Last");
            row5.createCell(1).setCellValue("Pilot First");
            row5.createCell(2).setCellValue("Business Name");
            row5.createCell(3).setCellValue("Operator Name"); //Unused Placeholder
            row5.createCell(4).setCellValue("Opr Fname"); //Unused Placeholder
            row5.createCell(5).setCellValue("A/C Reg");
            row5.createCell(6).setCellValue("Aircraft Model");
            row5.createCell(7).setCellValue("# Passes");
            row5.createCell(8).setCellValue("Business Address");
            row5.createCell(9).setCellValue("City");
            row5.createCell(10).setCellValue("State");
            row5.createCell(11).setCellValue("ZIP");
            row5.createCell(12).setCellValue("Phone");
            row5.createCell(13).setCellValue("Email");

            XSSFRow row6 = sh.createRow(sh.getLastRowNum() + 1);
            String [] pilot = textFieldPilot.getText().split(" ");
            String pilotFirst;
            String pilotLast;
            if(pilot.length == 2){
                pilotFirst = pilot[0];
                pilotLast = pilot[1];
            } else if(pilot.length == 3){
                pilotFirst = pilot[0];
                pilotLast = pilot[1] + " " + pilot[2];
            } else {
                pilotLast = "";
                pilotFirst = textFieldPilot.getText();
            }
            row6.createCell(0).setCellValue(pilotLast);
            row6.createCell(1).setCellValue(pilotFirst);
            row6.createCell(2).setCellValue(textFieldBusinessName.getText());
            //Unused Placeholder
            //Unused Placeholder
            row6.createCell(5).setCellValue(textFieldRegNum.getText());
            row6.createCell(6).setCellValue(comboBoxAircraftMake.getSelectionModel().getSelectedItem().toString() + " " + comboBoxAircraftModel.getSelectionModel().getSelectedItem().toString());
            int passes = 0;
            if(cBPass1.isSelected()){passes++;}
            if(cBPass2.isSelected()){passes++;}
            if(cBPass3.isSelected()){passes++;}
            if(cBPass4.isSelected()){passes++;}
            if(cBPass5.isSelected()){passes++;}
            if(cBPass6.isSelected()){passes++;}
            row6.createCell(7).setCellValue(passes);
            row6.createCell(8).setCellValue(textFieldStreet.getText());
            row6.createCell(9).setCellValue(textFieldCity.getText());
            row6.createCell(10).setCellValue(textFieldState.getText());
            row6.createCell(11).setCellValue(textFieldZIP.getText());
            row6.createCell(12).setCellValue(textFieldPhone.getText());
            row6.createCell(13).setCellValue(textFieldEmail.getText());

            sh.getRow(4).getCell(2).setCellValue(passes);

            for(int i=0; i<14; i++){
                sh.autoSizeColumn(i);
            }
            try {
                reportFile = new File(currentDirectory+"/SAFEReport "+currentDirectory.getName()+".xlsx");
                wb.write(new FileOutputStream(reportFile));
            } catch (IOException e) {
                e.printStackTrace();
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("File Save Error");
                alert.setHeaderText("Unable to Write Report File");
                alert.setContentText("Check if file is open in another program and try again.");
                alert.showAndWait();
            }
        } else {

            try {
                System.out.println("Not New");
                FileInputStream fis = new FileInputStream(reportFile);
                Workbook wb = WorkbookFactory.create(fis);
                Sheet sh = wb.getSheet("Fly-In Data");

                sh.getRow(0).getCell(2).setCellValue(header1);
                sh.getRow(1).getCell(2).setCellValue(header2);
                sh.getRow(2).getCell(2).setCellValue(header3);
                sh.getRow(3).getCell(2).setCellValue(header4);

                boolean alreadyLogged = false;
                int activeRow = 0;
                int passes = 0;
                for(int i=6; i<sh.getLastRowNum()+1; i++){
                    if(sh.getRow(i).getCell(5).getStringCellValue().equals(textFieldRegNum.getText())){
                        alreadyLogged = true;
                        activeRow = i;
                    }
                }
                if(!alreadyLogged) {
                    Row row = sh.createRow(sh.getLastRowNum() + 1);
                    String [] pilot = textFieldPilot.getText().split(" ");
                    String pilotFirst;
                    String pilotLast;
                    if(pilot.length == 2){
                        pilotFirst = pilot[0];
                        pilotLast = pilot[1];
                    } else if(pilot.length == 3){
                        pilotFirst = pilot[0];
                        pilotLast = pilot[1] + " " + pilot[2];
                    } else {
                        pilotLast = "";
                        pilotFirst = textFieldPilot.getText();
                    }
                    row.createCell(0).setCellValue(pilotLast);
                    row.createCell(1).setCellValue(pilotFirst);
                    row.createCell(2).setCellValue(textFieldBusinessName.getText());
                    //Unused Placeholder
                    //Unused Placeholder
                    row.createCell(5).setCellValue(textFieldRegNum.getText());
                    row.createCell(6).setCellValue(comboBoxAircraftMake.getSelectionModel().getSelectedItem().toString() + " " + comboBoxAircraftModel.getSelectionModel().getSelectedItem().toString());
                    if(cBPass1.isSelected()){passes++;}
                    if(cBPass2.isSelected()){passes++;}
                    if(cBPass3.isSelected()){passes++;}
                    if(cBPass4.isSelected()){passes++;}
                    if(cBPass5.isSelected()){passes++;}
                    if(cBPass6.isSelected()){passes++;}
                    row.createCell(7).setCellValue(passes);
                    row.createCell(8).setCellValue(textFieldStreet.getText());
                    row.createCell(9).setCellValue(textFieldCity.getText());
                    row.createCell(10).setCellValue(textFieldState.getText());
                    row.createCell(11).setCellValue(textFieldZIP.getText());
                    row.createCell(12).setCellValue(textFieldPhone.getText());
                    row.createCell(13).setCellValue(textFieldEmail.getText());
                } else {
                    Row row = sh.getRow(activeRow);
                    int initialPassNum = (int) row.getCell(7).getNumericCellValue();

                    if(cBPass1.isSelected()){passes++;}
                    if(cBPass2.isSelected()){passes++;}
                    if(cBPass3.isSelected()){passes++;}
                    if(cBPass4.isSelected()){passes++;}
                    if(cBPass5.isSelected()){passes++;}
                    if(cBPass6.isSelected()){passes++;}
                    int newPassNum = initialPassNum + passes;
                    row.getCell(7).setCellValue(newPassNum);
                }

                int initialTotalPasses = (int) sh.getRow(4).getCell(2).getNumericCellValue();
                sh.getRow(4).getCell(2).setCellValue(initialTotalPasses+passes);


                for(int i=0; i<14; i++){
                    sh.autoSizeColumn(i);
                }
                FileOutputStream fos = new FileOutputStream(reportFile);
                wb.write(fos);
                fos.close();
            } catch (Exception e) {
                e.printStackTrace();
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("File Save Error");
                alert.setHeaderText("Unable to Write Report File");
                alert.setContentText("Check if file is open in another program and try again.");
                alert.showAndWait();
            }
        }
    }                       //XLSX

    public void importParamsWorkbook() {
        try{
            FileInputStream fis = new FileInputStream(paramsWorkbook);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet("AccuPatt");

            //Import Spectrometer Settings
            if (sh.getRow(1) != null && sh.getRow(1).getCell(1) != null) {
                EXCITATION_WAVELENGTH = (int) sh.getRow(1).getCell(1).getNumericCellValue();
            }
            if (sh.getRow(2) != null && sh.getRow(2).getCell(1) != null) {
                TARGET_WAVELENGTH = (int) sh.getRow(2).getCell(1).getNumericCellValue();
            }
            if (sh.getRow(3) != null && sh.getRow(3).getCell(1) != null) {
                BOXCAR_WIDTH = (int) sh.getRow(3).getCell(1).getNumericCellValue();
            }
            if (sh.getRow(4) != null && sh.getRow(4).getCell(1) != null) {
                INTEGRATION_TIME = (int) sh.getRow(4).getCell(1).getNumericCellValue();
            }

            //Import String Drive Settings
            if (sh.getRow(9) != null && sh.getRow(9).getCell(1) != null) {
                METRIC = !sh.getRow(9).getCell(1).getBooleanCellValue();
            }
            if(METRIC){
                if (sh.getRow(10) != null && sh.getRow(10).getCell(2) != null) {
                    FLIGHTLINE_LENGTH = sh.getRow(10).getCell(2).getNumericCellValue();
                }
                if (sh.getRow(11) != null && sh.getRow(11).getCell(2) != null) {
                    SAMPLE_LENGTH = sh.getRow(11).getCell(2).getNumericCellValue();
                }
                if (sh.getRow(12) != null && sh.getRow(12).getCell(2) != null) {
                    STRING_VEL = sh.getRow(12).getCell(2).getNumericCellValue();
                }
            } else {
                if (sh.getRow(10) != null && sh.getRow(10).getCell(1) != null) {
                    FLIGHTLINE_LENGTH = sh.getRow(10).getCell(1).getNumericCellValue();
                }
                if (sh.getRow(11) != null && sh.getRow(11).getCell(1) != null) {
                    SAMPLE_LENGTH = sh.getRow(11).getCell(1).getNumericCellValue();
                }
                if (sh.getRow(12) != null && sh.getRow(12).getCell(1) != null) {
                    STRING_VEL = sh.getRow(12).getCell(1).getNumericCellValue();
                }
            }
            STRING_TIME = FLIGHTLINE_LENGTH / STRING_VEL;

            //Import Scanner Settings
            if (sh.getRow(1) != null && sh.getRow(1).getCell(5) != null) {
                useSpreadFactor = (int) sh.getRow(1).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(2) != null && sh.getRow(2).getCell(5) != null) {
                spreadA = sh.getRow(2).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(3) != null && sh.getRow(3).getCell(5) != null) {
                spreadB = sh.getRow(3).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(4) != null && sh.getRow(4).getCell(5) != null) {
                spreadC = sh.getRow(4).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(5) != null && sh.getRow(5).getCell(5) != null) {
                DEFAULT_THRESHOLD = (int) sh.getRow(5).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(6) != null && sh.getRow(6).getCell(5) != null) {
                CROP_SCALE = sh.getRow(6).getCell(5).getNumericCellValue();
            }
            if (sh.getRow(7) != null && sh.getRow(7).getCell(5) != null) {
                CARD_SPACING = (int) sh.getRow(7).getCell(5).getNumericCellValue();
            }

            //Import Other Settings

            if (sh.getRow(15) != null && sh.getRow(15).getCell(1) != null) {
                currentDirectory = new File (sh.getRow(15).getCell(1).getStringCellValue());
            }
            if (sh.getRow(16) != null && sh.getRow(16).getCell(1) != null) {
                useLogo = sh.getRow(16).getCell(1).getBooleanCellValue();
            }
            if (sh.getRow(17) != null && sh.getRow(17).getCell(1) != null) {
                logoFile = sh.getRow(17).getCell(1).getStringCellValue();
            }
            if (sh.getRow(18) != null && sh.getRow(18).getCell(1) != null) {
                invertPH = sh.getRow(18).getCell(1).getBooleanCellValue();
            }

            //Import Headers
            if (sh.getRow(21) != null && sh.getRow(21).getCell(0) != null) {
                header1 = sh.getRow(21).getCell(0).getStringCellValue();
            }
            if (sh.getRow(22) != null && sh.getRow(22).getCell(0) != null) {
                header2 = sh.getRow(22).getCell(0).getStringCellValue();
            }
            if (sh.getRow(23) != null && sh.getRow(23).getCell(0) != null) {
                header3 = sh.getRow(23).getCell(0).getStringCellValue();
            }
            if (sh.getRow(24) != null && sh.getRow(24).getCell(0) != null) {
                header4 = sh.getRow(24).getCell(0).getStringCellValue();
            }

            fis.close();
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    public void updateParamsWorkbook(){
        try {
            FileInputStream fis = new FileInputStream(paramsWorkbook);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("AccuPatt");
            sh.getRow(1).getCell(1).setCellValue(EXCITATION_WAVELENGTH);
            sh.getRow(2).getCell(1).setCellValue(TARGET_WAVELENGTH);
            sh.getRow(3).getCell(1).setCellValue(BOXCAR_WIDTH);
            sh.getRow(4).getCell(1).setCellValue(INTEGRATION_TIME);
            sh.getRow(9).createCell(1).setCellValue(!METRIC);
            sh.getRow(9).createCell(2).setCellValue(METRIC);
            if(METRIC){
                sh.getRow(10).createCell(2).setCellValue(FLIGHTLINE_LENGTH);
                sh.getRow(11).createCell(2).setCellValue(SAMPLE_LENGTH);
                sh.getRow(12).createCell(2).setCellValue(STRING_VEL);
            } else {
                sh.getRow(10).getCell(1).setCellValue(FLIGHTLINE_LENGTH);
                sh.getRow(11).getCell(1).setCellValue(SAMPLE_LENGTH);
                sh.getRow(12).getCell(1).setCellValue(STRING_VEL);
            }
            sh.getRow(1).getCell(5).setCellValue(useSpreadFactor);
            sh.getRow(2).getCell(5).setCellValue(spreadA);
            sh.getRow(3).getCell(5).setCellValue(spreadB);
            sh.getRow(4).getCell(5).setCellValue(spreadC);
            sh.getRow(5).getCell(5).setCellValue(DEFAULT_THRESHOLD);
            sh.getRow(6).getCell(5).setCellValue(CROP_SCALE);
            sh.getRow(7).getCell(5).setCellValue(CARD_SPACING);
            //Placeholder
            sh.getRow(16).getCell(1).setCellValue(useLogo);
            sh.getRow(17).getCell(1).setCellValue(logoFile);
            sh.getRow(18).getCell(1).setCellValue(invertPH);
            if (sh.getRow(21)!=null && sh.getRow(21).getCell(0)!=null) {
                sh.getRow(21).getCell(0).setCellValue(header1);
            } else {
                sh.createRow(21).createCell(0).setCellValue(header1);
            }
            if (sh.getRow(22)!=null && sh.getRow(22).getCell(0)!=null) {
                sh.getRow(22).getCell(0).setCellValue(header2);
            } else {
                sh.createRow(22).createCell(0).setCellValue(header2);
            }
            if (sh.getRow(23)!=null && sh.getRow(23).getCell(0)!=null) {
                sh.getRow(23).getCell(0).setCellValue(header3);
            } else {
                sh.createRow(23).createCell(0).setCellValue(header3);
            }
            if (sh.getRow(24)!=null && sh.getRow(24).getCell(0)!=null) {
                sh.getRow(24).getCell(0).setCellValue(header4);
            } else {
                sh.createRow(24).createCell(0).setCellValue(header4);
            }
            FileOutputStream fos = new FileOutputStream(paramsWorkbook);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}
    }

    private void updateDirectoryInParams(){
        try {
            FileInputStream fis = new FileInputStream(paramsWorkbook);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("AccuPatt");

            if (sh.getRow(15).getCell(1) !=null) {
                sh.getRow(15).getCell(1).setCellValue(currentDirectory.getAbsolutePath());
            } else {
                sh.getRow(15).createCell(1).setCellValue(currentDirectory.getAbsolutePath());
            }

            FileOutputStream fos = new FileOutputStream(paramsWorkbook);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}
    }

    public void generatePDF(File pdfFile, boolean stringPage, boolean cardPage){
        pdfDocument = new Document(PageSize.LETTER);

        //Append Trims
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sh = wb.getSheet("Pattern Data");
            sh.getRow(0).createCell(2).setCellValue(patternObject.getTrim1L());
            sh.getRow(0).createCell(4).setCellValue(patternObject.getTrim2L());
            sh.getRow(0).createCell(6).setCellValue(patternObject.getTrim3L());
            sh.getRow(0).createCell(8).setCellValue(patternObject.getTrim4L());
            sh.getRow(0).createCell(10).setCellValue(patternObject.getTrim5L());
            sh.getRow(0).createCell(12).setCellValue(patternObject.getTrim6L());
            sh.getRow(1).createCell(2).setCellValue(patternObject.getTrim1R());
            sh.getRow(1).createCell(4).setCellValue(patternObject.getTrim2R());
            sh.getRow(1).createCell(6).setCellValue(patternObject.getTrim3R());
            sh.getRow(1).createCell(8).setCellValue(patternObject.getTrim4R());
            sh.getRow(1).createCell(10).setCellValue(patternObject.getTrim5R());
            sh.getRow(1).createCell(12).setCellValue(patternObject.getTrim6R());
            sh.getRow(2).createCell(2).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(4).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(6).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(8).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(10).setCellValue(patternObject.getTrimVertical1());
            sh.getRow(2).createCell(12).setCellValue(patternObject.getTrimVertical1());
            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}

        try {
            PdfWriter writer = PdfWriter.getInstance(pdfDocument, new FileOutputStream(pdfFile.toString()));
            //writer.setPageEvent(new WatermarkPageEvent());
            pdfDocument.open();
            PdfContentByte cb0 = writer.getDirectContent();

            if (stringPage) {

                cb0.moveTo(20, 715);
                cb0.lineTo(175, 715);
                cb0.lineTo(175, 775);
                cb0.lineTo(20, 775);
                cb0.lineTo(20, 715);

                cb0.moveTo(20, 630);
                cb0.lineTo(580, 630);

                cb0.moveTo(455, 775);
                cb0.lineTo(455, 755);
                cb0.lineTo(580, 755);

                cb0.moveTo(20, 630);
                cb0.lineTo(20, 700);
                cb0.moveTo(35, 630);
                cb0.lineTo(35, 700);
                cb0.moveTo(140, 630);
                cb0.lineTo(140, 700);
                cb0.moveTo(155, 630);
                cb0.lineTo(155, 700);


                cb0.moveTo(410, 240);
                cb0.lineTo(550, 240);
                cb0.lineTo(550, 35);
                cb0.lineTo(410, 35);
                cb0.lineTo(410, 240);

                cb0.moveTo(410, 133);
                cb0.lineTo(420, 133);
                cb0.lineTo(418, 135);
                cb0.lineTo(420, 133);
                cb0.lineTo(418, 131);

                cb0.moveTo(550, 133);
                cb0.lineTo(540, 133);
                cb0.lineTo(542, 135);
                cb0.lineTo(540, 133);
                cb0.lineTo(542, 131);

                //Series
                cb0.moveTo(265, 630);
                cb0.lineTo(265, 700);
                cb0.moveTo(280, 630);
                cb0.lineTo(280, 700);

                //Model
                if (!cardPage) {
                    cb0.moveTo(475, 630);
                    cb0.lineTo(475, 715);
                    cb0.moveTo(490, 630);
                    cb0.lineTo(490, 715);
                } else {
                    cb0.moveTo(550,630);
                    cb0.lineTo(550,700);
                    cb0.moveTo(580,630);
                    cb0.lineTo(580,700);
                }

                cb0.stroke();

                //Thin Divider Lines
                cb0.setLineWidth(0.25);
                cb0.moveTo(35, 662);
                cb0.lineTo(140, 662);
                cb0.moveTo(155, 672);
                cb0.lineTo(265, 672);
                if(cardPage){
                    cb0.moveTo(520,630);
                    cb0.lineTo(520, 700);
                }
                cb0.stroke();

                //Color Coding
                cb0.setLineWidth(1);
                //Pass 1
                if (isScannedPass1) {
                    //Green
                    cb0.setRGBColorStroke(50, 132, 50);
                    cb0.moveTo(370, 702);
                    cb0.lineTo(387, 702);
                    cb0.stroke();
                }

                //Pass 2
                if (isScannedPass2) {
                    //Red
                    cb0.setRGBColorStroke(233, 150, 122);
                    cb0.moveTo(395, 702);
                    cb0.lineTo(412, 702);
                    cb0.stroke();
                }

                //Pass 3
                if (isScannedPass3) {
                    //Yellow
                    cb0.setRGBColorStroke(255, 204, 0);
                    cb0.moveTo(420, 702);
                    cb0.lineTo(437, 702);
                    cb0.stroke();
                }


                if(cardPage){
                    if(isScannedPass4) {
                        cb0.setLineDash(3,2,0);
                        cb0.setRGBColorStroke(50,132,50);
                        cb0.moveTo(445,702);
                        cb0.lineTo(462,702);
                        cb0.stroke();
                    }
                    if(isScannedPass5) {
                        cb0.setLineDash(3,2,0);
                        cb0.setRGBColorStroke(233,150,122);
                        cb0.moveTo(470,702);
                        cb0.lineTo(487,702);
                        cb0.stroke();
                    }
                    if(isScannedPass6) {
                        cb0.setLineDash(3,2,0);
                        cb0.setRGBColorStroke(255,204,0);
                        cb0.moveTo(495,702);
                        cb0.lineTo(512,702);
                        cb0.stroke();
                    }
                }

                cb0.beginText();

                //Business/Pilot Info
                try {
                    cb0.setFontAndSize(BaseFont.createFont(), 8);
                    cb0.setTextMatrix(25, 760);
                    cb0.showText(textFieldPilot.getText());
                    cb0.setTextMatrix(25, 750);
                    cb0.showText(textFieldBusinessName.getText());
                    cb0.setTextMatrix(25, 740);
                    cb0.showText(textFieldStreet.getText());
                    cb0.setTextMatrix(25, 730);
                    cb0.showText(textFieldCity.getText() + ", " + textFieldState.getText() + " " + textFieldZIP.getText());

                    cb0.setTextMatrix(25, 720);
                    if (textFieldPhone.getText().length() == 10) {
                        String p = textFieldPhone.getText();
                        cb0.showText("(" + p.substring(0, 3) + ") " + p.substring(3, 6) + "-" + p.substring(6, 10));
                    } else {
                        cb0.showText(textFieldPhone.getText());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Applicator Data and try again.");
                    alert.showAndWait();
                }


                //Headers
                try {
                    cb0.setFontAndSize(BaseFont.createFont(), 12);
                    if (header1 != null) {
                        cb0.setTextMatrix(200, 765);
                        cb0.showText(header1);
                    }
                    if (header2 != null) {
                        cb0.setTextMatrix(200, 750);
                        cb0.showText(header2);
                    }
                    if (header3 != null) {
                        cb0.setTextMatrix(200, 735);
                        cb0.showText(header3);
                    }
                    if (header4 != null) {
                        cb0.setTextMatrix(200, 720);
                        cb0.showText(header4);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Headers and try again.");
                    alert.showAndWait();
                }

                try {
                    //Title
                    cb0.setFontAndSize(BaseFont.createFont(), 22);
                    cb0.setTextMatrix(460, 760);
                    cb0.showText(textFieldRegNum.getText() + " - " + textFieldSeriesNum.getText());

                    //Aircraft Data
                    cb0.setFontAndSize(BaseFont.createFont(), 10);
                    cb0.setTextMatrix(0, 1, -1, 0, 30, 650);
                    cb0.showText("Aircraft");
                    cb0.setFontAndSize(BaseFont.createFont(), 8);
                    cb0.setTextMatrix(40, 695);
                    cb0.showText("Reg. #:");
                    cb0.setTextMatrix(75, 695);
                    cb0.showText(textFieldRegNum.getText());
                    cb0.setTextMatrix(40, 685);
                    cb0.showText("Series #:");
                    cb0.setTextMatrix(75, 685);
                    cb0.showText(textFieldSeriesNum.getText());
                    cb0.setTextMatrix(40, 675);
                    cb0.showText("Make:");
                    cb0.setTextMatrix(75, 675);
                    cb0.showText(comboBoxAircraftMake.getValue().toString());
                    cb0.setTextMatrix(40, 665);
                    cb0.showText("Model:");
                    cb0.setTextMatrix(75, 665);
                    cb0.showText(comboBoxAircraftModel.getValue().toString());
                    double wingSpan = 0;
                    double boomWidth = 0;
                    cb0.setTextMatrix(40, 655);
                    cb0.showText("% Boom Width:");
                    cb0.setTextMatrix(110, 655);
                    if (!textFieldWingSpan.getText().isEmpty() && !textFieldBoomWidth.getText().isEmpty()) {
                        wingSpan = Double.parseDouble(textFieldWingSpan.getText());
                        boomWidth = Double.parseDouble(textFieldBoomWidth.getText());
                        int percentBW = (int) Math.round((boomWidth / wingSpan) * 100);
                        cb0.showText(Integer.toString(percentBW));
                    } else {
                        cb0.showText("-");
                    }
                    cb0.setTextMatrix(40, 645);
                    cb0.showText(labelBoomDrop.getText());
                    cb0.setTextMatrix(110, 645);
                    if (!textFieldBoomDrop.getText().isEmpty()) {
                        cb0.showText(Integer.toString(Integer.parseInt(textFieldBoomDrop.getText())));
                    } else {
                        cb0.showText("-");
                    }
                    cb0.setTextMatrix(40, 635);
                    cb0.showText("Winglets?:");
                    cb0.setTextMatrix(110, 635);
                    cb0.showText(comboBoxWinglets.getValue().toString());
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Aircraft Data and try again.");
                    alert.showAndWait();
                }

                //Config Data
                try {
                    cb0.setFontAndSize(BaseFont.createFont(), 10);
                    cb0.setTextMatrix(0, 1, -1, 0, 150, 635);
                    cb0.showText("Configuration");
                    cb0.setFontAndSize(BaseFont.createFont(), 8);
                    cb0.setTextMatrix(160, 695);
                    cb0.showText(labelBoomPress.getText());
                    cb0.setTextMatrix(245, 695);
                    cb0.showText(textFieldBoomPressure.getText());
                    cb0.setTextMatrix(160, 685);
                    cb0.showText(labelTargetRate.getText());
                    cb0.setTextMatrix(245, 685);
                    cb0.showText(textFieldTargetRate.getText());
                    cb0.setTextMatrix(160, 675);
                    cb0.showText(labelTargetSwath.getText());
                    cb0.setTextMatrix(245, 675);
                    cb0.showText(textFieldTargetSwath.getText());
                    cb0.setTextMatrix(160, 665);
                    //cb0.showText("Set #1:");
                    //cb0.setTextMatrix(320, 695);
                    cb0.showText(comboBoxNozzle1Type.getValue().toString() + " (x" + textFieldNozzle1Quant.getText() + ")");
                    cb0.setTextMatrix(160, 655);
                    if (!comboBoxNozzle1Size.getValue().equals("") && !comboBoxNozzle1Def.getValue().equals("N/A")) {
                        cb0.showText("Orif. #: " + comboBoxNozzle1Size.getValue().toString() + ", Def. = " + Integer.toString((int) Double.parseDouble(comboBoxNozzle1Def.getValue().toString())) + "\u00b0");
                    } else {
                        cb0.showText("Orif. #: N/A, Def: N/A");
                    }

                    if (!comboBoxNozzle2Type.getValue().equals("Select Nozzle")) {
                        cb0.setTextMatrix(160, 645);
                        //cb0.showText("Set #2:");
                        //cb0.setTextMatrix(320, 675);
                        cb0.showText(comboBoxNozzle2Type.getValue().toString() + " (x" + textFieldNozzle2Quant.getText() + ")");
                        cb0.setTextMatrix(160, 635);
                        cb0.showText("Orif. #: " + comboBoxNozzle2Size.getValue().toString() + ", Def. = " + Integer.toString((int) Double.parseDouble(comboBoxNozzle2Def.getValue().toString())) + "\u00b0");
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Nozzle and Output Data and try again.");
                    alert.showAndWait();
                }

                //Series (1-3) and Model Data displayed if not printing card page
                if (!cardPage) {
                    //Series Data
                    try {
                        cb0.setFontAndSize(BaseFont.createFont(), 10);
                        cb0.setTextMatrix(0, 1, -1, 0, 275, 638);
                        cb0.showText("Series Data");
                        cb0.setFontAndSize(BaseFont.createFont(), 6);
                        cb0.setTextMatrix(370, 695);
                        cb0.showText("Pass 1");
                        cb0.setTextMatrix(395, 695);
                        cb0.showText("Pass 2");
                        cb0.setTextMatrix(420, 695);
                        cb0.showText("Pass 3");
                        cb0.setTextMatrix(445, 695);
                        cb0.showText("Avg");

                        cb0.setFontAndSize(BaseFont.createFont(), 8);
                        cb0.setTextMatrix(285, 685);
                        cb0.showText("Airspeed (" + uAirSpeed + "):");
                        cb0.setTextMatrix(285, 675);
                        cb0.showText("Release Height (" + uHeight + "):");
                        cb0.setTextMatrix(285, 665);
                        cb0.showText("Wind Velocity (" + uWindSpeed + "):");
                        cb0.setTextMatrix(285, 655);
                        cb0.showText("Cross-Wind (" + uWindSpeed + "):");
                        cb0.setTextMatrix(285, 645);
                        if (!METRIC) {
                            cb0.showText("Ambient Temp (F):");
                        } else {
                            cb0.showText("Ambient Temp (C):");
                        }
                        cb0.setTextMatrix(285, 635);
                        cb0.showText("Rel Humidity (%):");

                        if(isScannedPass1){
                            cb0.setTextMatrix(370, 685);
                            cb0.showText(Integer.toString(airspeedCalc(1)));
                            cb0.setTextMatrix(370, 675);
                            cb0.showText(textFieldSH1.getText());
                            cb0.setTextMatrix(370, 665);
                            cb0.showText(textFieldWV1.getText());
                            cb0.setTextMatrix(370, 655);
                            cb0.showText(Double.toString(crossWindCalc(1)));
                        } else {
                            cb0.setTextMatrix(370, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass2){
                            cb0.setTextMatrix(395, 685);
                            cb0.showText(Integer.toString(airspeedCalc(2)));
                            cb0.setTextMatrix(395, 675);
                            cb0.showText(textFieldSH2.getText());
                            cb0.setTextMatrix(395, 665);
                            cb0.showText(textFieldWV2.getText());
                            cb0.setTextMatrix(395, 655);
                            cb0.showText(Double.toString(crossWindCalc(2)));
                        } else {
                            cb0.setTextMatrix(395, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass3){
                            cb0.setTextMatrix(420, 685);
                            cb0.showText(Integer.toString(airspeedCalc(3)));
                            cb0.setTextMatrix(420, 675);
                            cb0.showText(textFieldSH3.getText());
                            cb0.setTextMatrix(420, 665);
                            cb0.showText(textFieldWV3.getText());
                            cb0.setTextMatrix(420, 655);
                            cb0.showText(Double.toString(crossWindCalc(3)));
                        } else {
                            cb0.setTextMatrix(420, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 655);
                            cb0.showText("-");
                        }

                        //Find Averages
                        cb0.setTextMatrix(445, 685);
                        cb0.showText(Integer.toString(averageI(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6), 1)));
                        cb0.setTextMatrix(445, 675);
                        cb0.showText(Double.toString(averageD(getSprayHeight(1),getSprayHeight(2),getSprayHeight(3),getSprayHeight(4),getSprayHeight(5),getSprayHeight(6),1)));
                        cb0.setTextMatrix(445, 665);
                        cb0.showText(Double.toString(averageD(getWindVel(1),getWindVel(2),getWindVel(3),getWindVel(4),getWindVel(5),getWindVel(6),1)));
                        cb0.setTextMatrix(445, 655);
                        cb0.showText(Double.toString(averageD(crossWindCalc(1),crossWindCalc(2),crossWindCalc(3),crossWindCalc(4),crossWindCalc(5),crossWindCalc(6),1)));

                        cb0.setTextMatrix(370, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(395, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(420, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(445, 645);
                        cb0.showText(textFieldAT.getText());
                        cb0.setTextMatrix(285, 635);
                        cb0.showText("Rel Humidity (%):");
                        cb0.setTextMatrix(370, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(395, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(420, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(445, 635);
                        cb0.showText(textFieldRH.getText());
                    } catch (Exception e) {
                        e.printStackTrace();
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setTitle("Error");
                        alert.setHeaderText("Unable to Generate Report");
                        alert.setContentText("Check Series Data and try again.");
                        alert.showAndWait();
                    }

                    if (!textAreaNotes.getText().isEmpty()) {
                        cb0.setTextMatrix(50, 620);
                        cb0.showText("Notes: " + textAreaNotes.getText());
                    }

                    //Model
                    try {
                        cb0.setFontAndSize(BaseFont.createFont(), 10);
                        cb0.setTextMatrix(0, 1, -1, 0, 485, 645);
                        cb0.showText("USDA Model");
                        cb0.setFontAndSize(BaseFont.createFont(), 8);
                        cb0.setTextMatrix(495, 705);
                        if (!METRIC) {
                            cb0.showText("Est. GPA:");
                        } else {
                            cb0.showText("Est. L/ha:");
                        }
                        cb0.setTextMatrix(550, 705);
                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
                            if (!METRIC) {
                                double estGPA = calcEstGPA();
                                if(estGPA!=-9) {
                                    cb0.showText(String.format("%.2f", estGPA));
                                } else {
                                    cb0.showText("Unknown");
                                }
                            } else {
                                double estLPHA = calcEstLPHA();
                                if(estLPHA!=-9) {
                                    cb0.showText(String.format("%.2f", estLPHA));
                                } else {
                                    cb0.showText("Unknown");
                                }
                            }
                        }
                        cb0.setTextMatrix(495, 695);
                        if (!METRIC) {
                            cb0.showText("Est. GPM:");
                        } else {
                            cb0.showText("Est. LPM:");
                        }
                        cb0.setTextMatrix(550, 695);
                        if (!textFieldNozzle1Quant.getText().isEmpty()) {
                            if (!METRIC) {
                                double estGPM = calcEstGPM();
                                if(estGPM!=-9) {
                                    cb0.showText(String.format("%.2f", estGPM));
                                } else {
                                    cb0.showText("Unknown");
                                }
                            } else {
                                double estLPM = calcEstLPM();
                                if(estLPM!=-9) {
                                    cb0.showText(String.format("%.2f", estLPM));
                                } else {
                                    cb0.showText("Unknown");
                                }
                            }
                        }
                        cb0.setTextMatrix(495, 685);
                        cb0.showText("DSC:");
                        cb0.setTextMatrix(550, 685);
                        cb0.showText(getDSC());
                        cb0.setTextMatrix(495, 675);
                        cb0.showText("DV0.1 (" + "\u00B5" + "m):");
                        cb0.setTextMatrix(550, 675);
//                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                            cb0.showText(Integer.toString(getDV01()));
//                        }
                        int dv1 = getDV01();
                        if(dv1!=-9){
                            cb0.showText(Integer.toString(dv1) + " \u00B5" + "m");
                        } else {
                            cb0.showText("Unknown");
                        }
                        cb0.setTextMatrix(495, 665);
                        cb0.showText("VMD (" + "\u00B5" + "m):");
                        cb0.setTextMatrix(550, 665);
//                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                            cb0.showText(Integer.toString(getDV05()));
//                        }
                        int dv5 = getDV05();
                        if(dv5!=-9){
                            cb0.showText(Integer.toString(dv5) + " \u00B5" + "m");
                        } else {
                            cb0.showText("Unknown");
                        }
                        cb0.setTextMatrix(495, 655);
                        cb0.showText("DV0.9 (" + "\u00B5" + "m):");
                        cb0.setTextMatrix(550, 655);
//                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                            cb0.showText(Integer.toString(getDV09()));
//                        }
                        int dv9 = getDV09();
                        if(dv9!=-9){
                            cb0.showText(Integer.toString(dv9) + " \u00B5" + "m");
                        } else {
                            cb0.showText("Unknown");
                        }
                        cb0.setTextMatrix(495, 645);
                        cb0.showText("RS:");
                        cb0.setTextMatrix(550, 645);
//                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                            cb0.showText(String.format("%.2f", getRS()));
//                        }
                        double rs = getRS();
                        if(rs!=-9 && rs!=0.00){
                            cb0.showText(String.format("%.2f", rs));
                        } else {
                            cb0.showText("Unknown");
                        }
                        cb0.setTextMatrix(495, 635);
                        cb0.showText("%<100" + "\u00B5" + "m:");
                        cb0.setTextMatrix(550, 635);
//                        if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                            cb0.showText(String.format("%.2f", getPercentLess100()));
//                        }
                        double pl100 = getPercentLess100();
                        if(pl100!=-9){
                            cb0.showText(String.format("%.2f", rs));
                        } else {
                            cb0.showText("Unknown");
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setTitle("Error");
                        alert.setHeaderText("Report Generation Error");
                        alert.setContentText("Unable to resolve Atomization Model calculations.");
                        alert.showAndWait();
                    }
                } else {
                    try{
                        //If printing card page, show series (1-6) data only
                        cb0.setFontAndSize(BaseFont.createFont(), 10);
                        cb0.setTextMatrix(0, 1, -1, 0, 275, 638);
                        cb0.showText("Series Data");
                        cb0.setFontAndSize(BaseFont.createFont(), 6);
                        cb0.setTextMatrix(370, 695);
                        cb0.showText("Pass 1");
                        cb0.setTextMatrix(395, 695);
                        cb0.showText("Pass 2");
                        cb0.setTextMatrix(420, 695);
                        cb0.showText("Pass 3");
                        cb0.setTextMatrix(445, 695);
                        cb0.showText("Pass 4");
                        cb0.setTextMatrix(470,695);
                        cb0.showText("Pass 5");
                        cb0.setTextMatrix(495,695);
                        cb0.showText("Pass 6");
                        cb0.setTextMatrix(525,695);
                        cb0.showText("Avg.");

                        cb0.setFontAndSize(BaseFont.createFont(), 8);
                        cb0.setTextMatrix(285, 685);
                        cb0.showText("Airspeed (" + uAirSpeed + "):");
                        cb0.setTextMatrix(285, 675);
                        cb0.showText("Release Height (" + uHeight + "):");
                        cb0.setTextMatrix(285, 665);
                        cb0.showText("Wind Velocity (" + uWindSpeed + "):");
                        cb0.setTextMatrix(285, 655);
                        cb0.showText("Cross-Wind (" + uWindSpeed + "):");
                        cb0.setTextMatrix(285, 645);
                        if (!METRIC) {
                            cb0.showText("Ambient Temp (F):");
                        } else {
                            cb0.showText("Ambient Temp (C):");
                        }
                        cb0.setTextMatrix(285, 635);
                        cb0.showText("Rel Humidity (%):");

                        if(isScannedPass1){
                            cb0.setTextMatrix(370, 685);
                            cb0.showText(Integer.toString(airspeedCalc(1)));
                            cb0.setTextMatrix(370, 675);
                            cb0.showText(textFieldSH1.getText());
                            cb0.setTextMatrix(370, 665);
                            cb0.showText(textFieldWV1.getText());
                            cb0.setTextMatrix(370, 655);
                            cb0.showText(Double.toString(crossWindCalc(1)));
                        } else {
                            cb0.setTextMatrix(370, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(370, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass2){
                            cb0.setTextMatrix(395, 685);
                            cb0.showText(Integer.toString(airspeedCalc(2)));
                            cb0.setTextMatrix(395, 675);
                            cb0.showText(textFieldSH2.getText());
                            cb0.setTextMatrix(395, 665);
                            cb0.showText(textFieldWV2.getText());
                            cb0.setTextMatrix(395, 655);
                            cb0.showText(Double.toString(crossWindCalc(2)));
                        } else {
                            cb0.setTextMatrix(395, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(395, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass3){
                            cb0.setTextMatrix(420, 685);
                            cb0.showText(Integer.toString(airspeedCalc(3)));
                            cb0.setTextMatrix(420, 675);
                            cb0.showText(textFieldSH3.getText());
                            cb0.setTextMatrix(420, 665);
                            cb0.showText(textFieldWV3.getText());
                            cb0.setTextMatrix(420, 655);
                            cb0.showText(Double.toString(crossWindCalc(3)));
                        } else {
                            cb0.setTextMatrix(420, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(420, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass4){
                            cb0.setTextMatrix(445, 685);
                            cb0.showText(Integer.toString(airspeedCalc(4)));
                            cb0.setTextMatrix(445, 675);
                            cb0.showText(textFieldSH4.getText());
                            cb0.setTextMatrix(445, 665);
                            cb0.showText(textFieldWV4.getText());
                            cb0.setTextMatrix(445, 655);
                            cb0.showText(Double.toString(crossWindCalc(4)));
                        } else {
                            cb0.setTextMatrix(445, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(445, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(445, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(445, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass5){
                            cb0.setTextMatrix(470, 685);
                            cb0.showText(Integer.toString(airspeedCalc(5)));
                            cb0.setTextMatrix(470, 675);
                            cb0.showText(textFieldSH5.getText());
                            cb0.setTextMatrix(470, 665);
                            cb0.showText(textFieldWV5.getText());
                            cb0.setTextMatrix(470, 655);
                            cb0.showText(Double.toString(crossWindCalc(5)));
                        } else {
                            cb0.setTextMatrix(470, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(470, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(470, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(470, 655);
                            cb0.showText("-");
                        }

                        if(isScannedPass6){
                            cb0.setTextMatrix(495, 685);
                            cb0.showText(Integer.toString(airspeedCalc(6)));
                            cb0.setTextMatrix(495, 675);
                            cb0.showText(textFieldSH6.getText());
                            cb0.setTextMatrix(495, 665);
                            cb0.showText(textFieldWV6.getText());
                            cb0.setTextMatrix(495, 655);
                            cb0.showText(Double.toString(crossWindCalc(6)));
                        } else {
                            cb0.setTextMatrix(495, 685);
                            cb0.showText("-");
                            cb0.setTextMatrix(495, 675);
                            cb0.showText("-");
                            cb0.setTextMatrix(495, 665);
                            cb0.showText("-");
                            cb0.setTextMatrix(495, 655);
                            cb0.showText("-");
                        }

                        //Find Averages
                        cb0.setTextMatrix(525, 685);
                        cb0.showText(Integer.toString(averageI(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6), 1)));
                        cb0.setTextMatrix(525, 675);
                        cb0.showText(Double.toString(averageD(getSprayHeight(1),getSprayHeight(2),getSprayHeight(3),getSprayHeight(4),getSprayHeight(5),getSprayHeight(6),1)));
                        cb0.setTextMatrix(525, 665);
                        cb0.showText(Double.toString(averageD(getWindVel(1),getWindVel(2),getWindVel(3),getWindVel(4),getWindVel(5),getWindVel(6),1)));
                        cb0.setTextMatrix(525, 655);
                        cb0.showText(Double.toString(averageD(crossWindCalc(1),crossWindCalc(2),crossWindCalc(3),crossWindCalc(4),crossWindCalc(5),crossWindCalc(6),1)));

                        cb0.setTextMatrix(370, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(395, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(420, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(445, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(470, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(495, 645);
                        cb0.showText("-");
                        cb0.setTextMatrix(525, 645);
                        cb0.showText(textFieldAT.getText());


                        cb0.setTextMatrix(370, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(395, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(420, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(445, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(470, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(495, 635);
                        cb0.showText("-");
                        cb0.setTextMatrix(525, 635);
                        cb0.showText(textFieldRH.getText());

                        cb0.setFontAndSize(BaseFont.createFont(), 10);
                        cb0.setTextMatrix(0, 1, -1, 0, 562, 640);
                        cb0.showText("Droplet Data");
                        cb0.setTextMatrix(0,1,-1,0,575,645);
                        cb0.showText("on Page 2");
                    } catch (Exception e) {
                        e.printStackTrace();
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setTitle("Error");
                        alert.setHeaderText("Unable to Generate Report");
                        alert.setContentText("Check Series Data and try again.");
                        alert.showAndWait();
                    }

                    if (!textAreaNotes.getText().isEmpty()) {
                        cb0.setTextMatrix(50, 620);
                        cb0.showText("Notes: " + textAreaNotes.getText());
                    }
                }

                cb0.setFontAndSize(BaseFont.createFont(), 12);
                cb0.setTextMatrix(0, 1, -1, 0, 40, 500);
                cb0.showText("Comparison");
                cb0.setTextMatrix(0, 1, -1, 0, 40, 330);
                cb0.showText("Average");
                cb0.setFontAndSize(BaseFont.createFont(), 10);
                cb0.setTextMatrix(0, 1, -1, 0, 40, 170);
                cb0.showText("Race Track");
                cb0.setTextMatrix(0, 1, -1, 0, 40, 50);
                cb0.showText("Back and Forth");

                double swath = Double.parseDouble(labelSwathFinal.getText());
                cb0.setFontAndSize(BaseFont.createFont(), 8);
                cb0.setTextMatrix(420, 230);
                if (!METRIC) {
                    cb0.showText("Swath (FT)");
                } else {
                    cb0.showText("Swath (m)");
                }
                cb0.setTextMatrix(470, 230);
                cb0.showText("CV (RT)");
                cb0.setTextMatrix(510, 230);
                cb0.showText("CV (BF)");

                for (int i = 0; i < 19; i++) {
                    if(METRIC){
                        double sw = swath + (-4.5 + (i*0.5));
                        cb0.setTextMatrix(435, 220 - (i * 10));
                        cb0.showText(String.format("%.1f",sw));
                        cb0.setTextMatrix(477, 220 - (i * 10));
                        if (calcRTCV(sw)>0) {
                            cb0.showText(Integer.toString(calcRTCV(sw)) + " %");
                        } else {
                            cb0.showText("N/A");
                        }
                        cb0.setTextMatrix(517, 220 - (i * 10));
                        if (calcBFCV(sw)>0) {
                            cb0.showText(Integer.toString(calcBFCV(sw)) + " %");
                        } else {
                            cb0.showText("N/A");
                        }
                    } else {
                        double sw = swath + (-9+i);
                        cb0.setTextMatrix(435, 220 - (i * 10));
                        cb0.showText(Integer.toString((int) sw));
                        cb0.setTextMatrix(477, 220 - (i * 10));
                        if (calcRTCV(sw)>0) {
                            cb0.showText(Integer.toString(calcRTCV(sw)) + " %");
                        } else {
                            cb0.showText("N/A");
                        }
                        cb0.setTextMatrix(517, 220 - (i * 10));
                        if (calcBFCV(sw)>0) {
                            cb0.showText(Integer.toString(calcBFCV(sw)) + " %");
                        } else {
                            cb0.showText("N/A");
                        }
                    }

                }

                cb0.endText();

                try {
                    PdfContentByte cb01 = writer.getDirectContent();

                    tabPaneMain.getSelectionModel().select(1);

                    WritableImage wim11 = lineChartSeriesOverlay.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput11 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim11, null), "png", byteOutput11);
                    Image aCSeriesOverlay = Image.getInstance(byteOutput11.toByteArray());
                    aCSeriesOverlay.scaleAbsolute(500, 175);
                    cb01.addImage(aCSeriesOverlay, 500, 0, 0, 175, 50, 430);

                    WritableImage wim12 = areaChartSeriesAverage.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput12 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim12, null), "png", byteOutput12);
                    Image aCSeriesAverage = Image.getInstance(byteOutput12.toByteArray());
                    aCSeriesAverage.scaleAbsolute(500, 175);
                    cb01.addImage(aCSeriesAverage, 500, 0, 0, 175, 50, 250);

                    tabPaneMain.getSelectionModel().select(2);

                    WritableImage wim14 = areaChartRacetrack.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput14 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim14, null), "png", byteOutput14);
                    Image aCRacetrack = Image.getInstance(byteOutput14.toByteArray());
                    aCRacetrack.scaleAbsolute(350, 110);
                    cb01.addImage(aCRacetrack, 350, 0, 0, 115, 50, 135);

                    WritableImage wim15 = areaChartBackAndForth.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput15 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim15, null), "png", byteOutput15);
                    Image aCBackAndForth = Image.getInstance(byteOutput15.toByteArray());
                    aCBackAndForth.scaleAbsolute(350, 110);
                    cb01.addImage(aCBackAndForth, 350, 0, 0, 115, 50, 20);

                    cb0.beginText();
                    cb0.setFontAndSize(BaseFont.createFont(), 6);
                    cb0.setTextMatrix(85, 580);
                    cb0.showText("Left Wing");
                    cb0.setTextMatrix(505, 580);
                    cb0.showText("Right Wing");
                    cb0.setTextMatrix(85, 400);
                    cb0.showText("Left Wing");
                    cb0.setTextMatrix(505, 400);
                    cb0.showText("Right Wing");
                    cb0.endText();

                    cb0.beginText();
                    cb0.setFontAndSize(BaseFont.createFont(), 4);
                    cb0.setTextMatrix(70, 240);
                    cb0.showText("Yellow: L1 into page");
                    cb0.setTextMatrix(70, 235);
                    cb0.showText("Green: C into page");
                    cb0.setTextMatrix(70, 230);
                    cb0.showText("Red: R1 into page");

                    cb0.setTextMatrix(70, 125);
                    cb0.showText("Red: L1 out of page");
                    cb0.setTextMatrix(70, 120);
                    cb0.showText("Green: C into page");
                    cb0.setTextMatrix(70, 115);
                    cb0.showText("Yellow: R1 out of page");
                    cb0.endText();

                    if (useLogo) {
                        try {
                            Image logo = Image.getInstance(logoFile);
                            logo.setAbsolutePosition(370, 710);
                            logo.scaleToFit(75, 75);
                            pdfDocument.add(logo);
                        } catch (IOException e) {
                            e.printStackTrace();
                            Alert alert = new Alert(Alert.AlertType.ERROR, "Could not find logo file at specified location");
                            alert.showAndWait();
                        }
                    }


                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Plot or Logo images not found.");
                    alert.showAndWait();
                }
            }

            if (cardPage) {
                //NewPage----------------------------------------------------------------------------------------------
                if (stringPage) {
                    pdfDocument.newPage();
                }

                PdfContentByte cbRect = writer.getDirectContent();
                Rectangle rectangle = new Rectangle(115,680,585,700);
                rectangle.setBackgroundColor(BaseColor.LIGHT_GRAY);
                cbRect.rectangle(rectangle);

                //Title - Upper Right Hand Corner
                cb0.moveTo(455, 775);
                cb0.lineTo(455, 755);
                cb0.lineTo(580, 755);
                //Business/Pilot Info
                cb0.moveTo(20, 715);
                cb0.lineTo(175, 715);
                cb0.lineTo(175, 775);
                cb0.lineTo(20, 775);
                cb0.lineTo(20, 715);
                //Header Divider
                cb0.moveTo(20,680);
                cb0.lineTo(20, 660);
                cb0.lineTo(585, 660);
                cb0.lineTo(585,700);
                cb0.moveTo(20, 680);
                cb0.lineTo(585, 680);
                cb0.moveTo(115, 680);
                cb0.lineTo(115, 700);
                cb0.lineTo(585,700);

                cb0.stroke();

                //Thin Dividers
                cb0.setLineWidth(0.25);
                cb0.moveTo(115,660);
                cb0.lineTo(115,680);
                cb0.moveTo(250,660);
                cb0.lineTo(250,710);
                cb0.moveTo(300, 660);
                cb0.lineTo(300,710);
                cb0.moveTo(350,660);
                cb0.lineTo(350, 710);
                cb0.moveTo(400,660);
                cb0.lineTo(400,710);
                cb0.moveTo(465,660);
                cb0.lineTo(465,710);
                cb0.stroke();


                //Title - Upper Right Hand Corner
                cb0.beginText();
                cb0.setFontAndSize(BaseFont.createFont(), 22);
                cb0.setTextMatrix(460, 760);
                cb0.showText(textFieldRegNum.getText() + " - " + textFieldSeriesNum.getText());

                //Business/Pilot Info
                try {
                    cb0.setFontAndSize(BaseFont.createFont(), 8);
                    cb0.setTextMatrix(25, 760);
                    cb0.showText(textFieldPilot.getText());
                    cb0.setTextMatrix(25, 750);
                    cb0.showText(textFieldBusinessName.getText());
                    cb0.setTextMatrix(25, 740);
                    cb0.showText(textFieldStreet.getText());
                    cb0.setTextMatrix(25, 730);
                    cb0.showText(textFieldCity.getText() + ", " + textFieldState.getText() + " " + textFieldZIP.getText());
                    cb0.setTextMatrix(25, 720);
                    if (textFieldPhone.getText().length() == 10) {
                        String p = textFieldPhone.getText();
                        cb0.showText("(" + p.substring(0, 3) + ") " + p.substring(3, 6) + "-" + p.substring(6, 10));
                    } else {
                        cb0.showText(textFieldPhone.getText());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Applicator Data and try again.");
                    alert.showAndWait();
                }

                //Headers
                try {
                    cb0.setFontAndSize(BaseFont.createFont(), 12);
                    if (header1 != null) {
                        cb0.setTextMatrix(200, 765);
                        cb0.showText(header1);
                    }
                    if (header2 != null) {
                        cb0.setTextMatrix(200, 750);
                        cb0.showText(header2);
                    }
                    if (header3 != null) {
                        cb0.setTextMatrix(200, 735);
                        cb0.showText(header3);
                    }
                    if (header4 != null) {
                        cb0.setTextMatrix(200, 720);
                        cb0.showText(header4);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Check Headers and try again.");
                    alert.showAndWait();
                }



                //Composite Data Title
                cb0.setFontAndSize(BaseFont.createFont(), 11);
                cb0.setTextMatrix(25, 665);
                cb0.showText("Composite Data");

                cb0.setTextMatrix(125, 685);
                cb0.showText("USDA Model Predicted");

                cb0.setFontAndSize(BaseFont.createFont(), 8);
                cb0.setTextMatrix(265, 705);
                cb0.showText("DV0.1");

                cb0.setTextMatrix(315, 705);
                cb0.showText("VMD");

                cb0.setTextMatrix(365,705);
                cb0.showText("DV0.9");

                cb0.setTextMatrix(407, 705);
                cb0.showText("Relative Span");

                cb0.setTextMatrix(475,705);
                cb0.showText("Droplet Spectrum Category");

                cb0.setFontAndSize(BaseFont.createFont(), 10);
                cb0.setTextMatrix(140, 665);
                cb0.showText(labelCCov.getText() + " Coverage");
                cb0.setTextMatrix(255, 665);
                cb0.showText(labelCDV01.getText());
                cb0.setTextMatrix(305, 665);
                cb0.showText(labelCVMD.getText());
                cb0.setTextMatrix(355, 665);
                cb0.showText(labelCDV09.getText());
                cb0.setTextMatrix(420, 665);
                cb0.showText(labelCRS.getText());

                //Translate abbreviation to full DSC
                String CDSC = labelCDSC.getText();
                if(CDSC.equalsIgnoreCase("XF")){
                    CDSC = "Extremely Fine";
                } else if (CDSC.equalsIgnoreCase("VF")){
                    CDSC = "Very Fine";
                } else if (CDSC.equalsIgnoreCase("F")){
                    CDSC = "Fine";
                } else if (CDSC.equalsIgnoreCase("M")){
                    CDSC = "Medium";
                } else if (CDSC.equalsIgnoreCase("C")){
                    CDSC = "Course";
                } else if (CDSC.equalsIgnoreCase("VC")){
                    CDSC = "Very Course";
                } else if (CDSC.equalsIgnoreCase("XC")){
                    CDSC = "Extremely Course";
                } else if (CDSC.equalsIgnoreCase("UC")){
                    CDSC = "Ultra Course";
                }
                cb0.setTextMatrix(485,665);
                cb0.showText(CDSC);

                try {
                    cb0.setTextMatrix(255,685);
//                    if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                        cb0.showText(Integer.toString(getDV01()) + " \u00B5" + "m");
//                    }
                    int dv1 = getDV01();
                    if(dv1!=-9){
                        cb0.showText(Integer.toString(dv1) + " \u00B5" + "m");
                    } else {
                        cb0.showText("Unknown");
                    }
                    cb0.setTextMatrix(305, 685);
//                    if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                        cb0.showText(Integer.toString(getDV05()) + " \u00B5" + "m");
//                    }
                    int dv5 = getDV05();
                    if(dv5!=-9){
                        cb0.showText(Integer.toString(dv5) + " \u00B5" + "m");
                    } else {
                        cb0.showText("Unknown");
                    }
                    cb0.setTextMatrix(355, 685);
//                    if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                        cb0.showText(Integer.toString(getDV09()) + " \u00B5" + "m");
//                    }
                    int dv9 = getDV09();
                    if(dv9!=-9){
                        cb0.showText(Integer.toString(dv9) + " \u00B5" + "m");
                    } else {
                        cb0.showText("Unknown");
                    }
                    cb0.setTextMatrix(420,685);
//                    if (!textFieldNozzle1Quant.getText().isEmpty() && (!textFieldGS1.getText().isEmpty() || !textFieldGS2.getText().isEmpty() || !textFieldGS3.getText().isEmpty())) {
//                        cb0.showText(String.format("%.2f", getRS()));
//                    }
                    double rs = getRS();
                    if(rs!=-9 && rs!=0.00){
                        cb0.showText(String.format("%.2f", rs));
                    } else {
                        cb0.setTextMatrix(410, 685);
                        cb0.showText("Unknown");
                    }
                    //Model convert abbreviation to full
                    CDSC = getDSC();
                    if(CDSC.equalsIgnoreCase("XF")){
                        CDSC = "Extremely Fine";
                    } else if (CDSC.equalsIgnoreCase("VF")){
                        CDSC = "Very Fine";
                    } else if (CDSC.equalsIgnoreCase("F")){
                        CDSC = "Fine";
                    } else if (CDSC.equalsIgnoreCase("M")){
                        CDSC = "Medium";
                    } else if (CDSC.equalsIgnoreCase("C")){
                        CDSC = "Course";
                    } else if (CDSC.equalsIgnoreCase("VC")){
                        CDSC = "Very Course";
                    } else if (CDSC.equalsIgnoreCase("XC")){
                        CDSC = "Extremely Course";
                    } else if (CDSC.equalsIgnoreCase("UC")){
                        CDSC = "Ultra Course";
                    } else {
                        CDSC = "Unknown";
                    }
                    cb0.setTextMatrix(485,685);
                    cb0.showText(CDSC);
                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Report Generation Error");
                    alert.setContentText("Unable to resolve Atomization Model calculations.");
                    alert.showAndWait();
                }

                //Card Data
                //cb0.setFontAndSize(BaseFont.createFont(),8);

                ArrayList<Boolean> table = new ArrayList<>();
                table.add(togL32.isSelected());
                table.add(togL24.isSelected());
                table.add(togL16.isSelected());
                table.add(togL8.isSelected());
                table.add(tog0.isSelected());
                table.add(togR8.isSelected());
                table.add(togR16.isSelected());
                table.add(togR24.isSelected());
                table.add(togR32.isSelected());

                int indexOfROI = 0;
                for (int i = 0; i < cardObject.getCardList().size(); i++) {
                    if(cardObject.getCardList().get(i)){
                        if(cardObject.getUsingCardList().get(i)){
                            cb0.setFontAndSize(BaseFont.createFont(), 10);
                            if (cardObject.getCardNameList().get(indexOfROI).length() > 4) {
                                cb0.setTextMatrix(23 + 65 * i, 480);
                            } else if (cardObject.getCardNameList().get(indexOfROI).length() == 4) {
                                cb0.setTextMatrix(35 + 65 * i, 480);
                            } else {
                                cb0.setTextMatrix(32 + 65 * i, 480);
                            }
                            cb0.showText(cardObject.getCardNameList().get(indexOfROI));
                            cb0.setFontAndSize(BaseFont.createFont(), 8);
                            cb0.setTextMatrix(25 + 65 * i, 470);
                            cb0.showText("VMD: " + Integer.toString(cardObject.getdV05s().get(indexOfROI).intValue()));
                            cb0.setTextMatrix(25 + 65 * i, 460);
                            cb0.showText("DV01: " + Integer.toString(cardObject.getdV01s().get(indexOfROI).intValue()));
                            cb0.setTextMatrix(25 + 65 * i, 450);
                            cb0.showText("DV09: " + Integer.toString(cardObject.getdV09s().get(indexOfROI).intValue()));
                            cb0.setTextMatrix(22 + 65 * i, 440);
                            cb0.showText("COV: " + String.format("%.2f", cardObject.getPercentCoverage().get(indexOfROI)) + "%");
                            cb0.setTextMatrix(21 + 65 * i, 430);
                            cb0.showText(Integer.toString(cardObject.getDropsPerSquareInch().get(indexOfROI).intValue()) + " drops/in^2");
                        }
                        indexOfROI++;
                    }

                }

                cb0.setTextMatrix(450, 735);
                cb0.showText("*See Page 1 - Pass " + spinnerCardPassNum.getValueFactory().getValue() + " for weather data.");

                if (useSpreadFactor == 2) {
                    cb0.setTextMatrix(430, 725);
                    cb0.showText("**DD = "+String.format("%.4f",spreadA)+"xDS^2 + "+String.format("%.4f",spreadB)+"xDS + "+String.format("%.4f",spreadC));
                } else {
                    cb0.setTextMatrix(390, 725);
                    cb0.showText("**DD = DS / (("+String.format("%.4f",spreadA)+" x DS^2) + ("+String.format("%.4f",spreadB)+" x DS) + "+String.format("%.4f",spreadC)+"))");
                }

                cb0.endText();



                try {
                    PdfContentByte cb01 = writer.getDirectContent();

                    tabPaneMain.getSelectionModel().select(3);

                    if (table.get(0)) {
                        WritableImage wim10 = imageViewCard0.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput10 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim10, null), "png", byteOutput10);
                        Image card0 = Image.getInstance(byteOutput10.toByteArray());
                        card0.scaleToFit(60, 160);
                        card0.setAbsolutePosition(15, 490);
                        cb01.addImage(card0);
                    }

                    if (table.get(1)) {
                        WritableImage wim11 = imageViewCard1.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput11 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim11, null), "png", byteOutput11);
                        Image card1 = Image.getInstance(byteOutput11.toByteArray());
                        card1.scaleToFit(60, 160);
                        card1.setAbsolutePosition(80, 490);
                        cb01.addImage(card1);
                    }

                    if (table.get(2)) {
                        WritableImage wim12 = imageViewCard2.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput12 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim12, null), "png", byteOutput12);
                        Image card2 = Image.getInstance(byteOutput12.toByteArray());
                        card2.scaleToFit(60, 160);
                        card2.setAbsolutePosition(145, 490);
                        cb01.addImage(card2);
                    }

                    if (table.get(3)) {
                        WritableImage wim13 = imageViewCard3.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput13 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim13, null), "png", byteOutput13);
                        Image card3 = Image.getInstance(byteOutput13.toByteArray());
                        card3.scaleToFit(60, 160);
                        card3.setAbsolutePosition(210, 490);
                        cb01.addImage(card3);
                    }

                    if (table.get(4)) {
                        WritableImage wim14 = imageViewCard4.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput14 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim14, null), "png", byteOutput14);
                        Image card4 = Image.getInstance(byteOutput14.toByteArray());
                        card4.scaleToFit(60, 160);
                        card4.setAbsolutePosition(275, 490);
                        cb01.addImage(card4);
                    }

                    if (table.get(5)) {
                        WritableImage wim15 = imageViewCard5.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput15 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim15, null), "png", byteOutput15);
                        Image card5 = Image.getInstance(byteOutput15.toByteArray());
                        card5.scaleToFit(60, 160);
                        card5.setAbsolutePosition(340, 490);
                        cb01.addImage(card5);
                    }

                    if (table.get(6)) {
                        WritableImage wim16 = imageViewCard6.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput16 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim16, null), "png", byteOutput16);
                        Image card6 = Image.getInstance(byteOutput16.toByteArray());
                        card6.scaleToFit(60, 160);
                        card6.setAbsolutePosition(405, 490);
                        cb01.addImage(card6);
                    }

                    if (table.get(7)) {
                        WritableImage wim17 = imageViewCard7.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput17 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim17, null), "png", byteOutput17);
                        Image card7 = Image.getInstance(byteOutput17.toByteArray());
                        card7.scaleToFit(60, 160);
                        card7.setAbsolutePosition(470, 490);
                        cb01.addImage(card7);
                    }

                    if (table.get(8)) {
                        WritableImage wim18 = imageViewCard8.snapshot(new SnapshotParameters(), null);
                        ByteArrayOutputStream byteOutput18 = new ByteArrayOutputStream();
                        ImageIO.write(SwingFXUtils.fromFXImage(wim18, null), "png", byteOutput18);
                        Image card8 = Image.getInstance(byteOutput18.toByteArray());
                        card8.scaleToFit(60, 160);
                        card8.setAbsolutePosition(535, 490);
                        cb01.addImage(card8);
                    }

                    tabPaneMain.getSelectionModel().select(4);

                    WritableImage wim20 = areaChartCoverage.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput20 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim20, null), "png", byteOutput20);
                    Image plotCoverage = Image.getInstance(byteOutput20.toByteArray());
                    plotCoverage.scaleToFit(550, 250);
                    plotCoverage.setAbsolutePosition(25, 230);
                    cb01.addImage(plotCoverage);

                    WritableImage wim21 = barChartDropletVolume.snapshot(new SnapshotParameters(), null);
                    ByteArrayOutputStream byteOutput21 = new ByteArrayOutputStream();
                    ImageIO.write(SwingFXUtils.fromFXImage(wim21, null), "png", byteOutput21);
                    Image plotDroplets = Image.getInstance(byteOutput21.toByteArray());
                    plotDroplets.scaleToFit(550, 250);
                    plotDroplets.setAbsolutePosition(25, 25);
                    cb01.addImage(plotDroplets);


                } catch (Exception e) {
                    e.printStackTrace();
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setTitle("Error");
                    alert.setHeaderText("Unable to Generate Report");
                    alert.setContentText("Card images or Plots not found.");
                    alert.showAndWait();
                }
            }

            pdfDocument.close();

            try {
                if (pdfFile != null) {
                    new Thread(new Runnable() {
                        @Override
                        public void run() {
                            try {
                                Desktop.getDesktop().open(pdfFile);
                            } catch (IOException e) {
                                e.printStackTrace();
                                Alert alert = new Alert(Alert.AlertType.ERROR);
                                alert.setTitle("Error");
                                alert.setHeaderText("Unable to Generate Report");
                                alert.setContentText("PDF Display Fail.");
                                alert.showAndWait();
                            }
                        }
                    }).start();
                }
            } catch (Exception e) {
                e.printStackTrace();
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle("Error");
                alert.setHeaderText("Unable to Generate Report");
                alert.setContentText("PDF Display Fail.");
                alert.showAndWait();
            }
        } catch (Exception e) {
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Error");
            alert.setHeaderText("Unable to Generate Report");
            alert.setContentText("File may already be open");
            alert.showAndWait();
        }
    }

    //Visibility and Clearing
    public void clearAircraftAndSeriesData(){
        textFieldPilot.clear();
        textFieldBusinessName.clear();
        textFieldPhone.clear();
        textFieldStreet.clear();
        textFieldCity.clear();
        textFieldState.clear();
        textFieldZIP.clear();
        textFieldEmail.clear();
        textFieldRegNum.clear();
        textFieldSeriesNum.clear();
        comboBoxAircraftMake.setValue("");
        comboBoxAircraftModel.setValue("");
        comboBoxAircraftModel.getItems().clear();
        comboBoxNozzle1Type.setValue("Select Nozzle");
        textFieldNozzle1Quant.clear();
        comboBoxNozzle1Size.setValue("N/A");
        comboBoxNozzle1Def.setValue("N/A");
        comboBoxNozzle2Type.setValue("Select Nozzle");
        textFieldNozzle2Quant.clear();
        comboBoxNozzle2Size.setValue("N/A");
        comboBoxNozzle2Def.setValue("N/A");
        textFieldBoomPressure.clear();
        textFieldTargetRate.clear();
        textFieldTargetSwath.clear();

        textFieldGS1.clear();
        textFieldGS2.clear();
        textFieldGS3.clear();
        textFieldGS4.clear();
        textFieldGS5.clear();
        textFieldGS6.clear();
        textFieldSH1.clear();
        textFieldSH2.clear();
        textFieldSH3.clear();
        textFieldSH4.clear();
        textFieldSH5.clear();
        textFieldSH6.clear();
        textFieldPH1.clear();
        textFieldPH2.clear();
        textFieldPH3.clear();
        textFieldPH4.clear();
        textFieldPH5.clear();
        textFieldPH6.clear();
        textFieldWD1.clear();
        textFieldWD2.clear();
        textFieldWD3.clear();
        textFieldWD4.clear();
        textFieldWD5.clear();
        textFieldWD6.clear();
        textFieldWV1.clear();
        textFieldWV2.clear();
        textFieldWV3.clear();
        textFieldWV4.clear();
        textFieldWV5.clear();
        textFieldWV6.clear();
        textFieldAT.clear();
        textFieldRH.clear();
        //Clear Extra Info
        textFieldTime.clear();
        textFieldWingSpan.clear();
        textFieldBoomWidth.clear();
        textFieldBoomDrop.clear();
        textFieldNozzleSpacing.clear();
        comboBoxWinglets.setValue("No");
        textAreaNotes.clear();
    }

    public void clearPatternData(){
        isScannedPass1 = false;
        isScannedPass2 = false;
        isScannedPass3 = false;
        isScannedPass4 = false;
        isScannedPass5 = false;
        isScannedPass6 = false;
        toggleButtonPass1.setDisable(true);
        toggleButtonPass2.setDisable(true);
        toggleButtonPass3.setDisable(true);
        toggleButtonPass4.setDisable(true);
        toggleButtonPass5.setDisable(true);
        toggleButtonPass6.setDisable(true);
        cBPass1.setSelected(false);
        cBPass2.setSelected(false);
        cBPass3.setSelected(false);
        cBPass4.setSelected(false);
        cBPass5.setSelected(false);
        cBPass6.setSelected(false);
        spinnerTrimL.getValueFactory().setValue(0);
        spinnerTrimR.getValueFactory().setValue(0);
        activePasses.clear();
        labelSwathFinal.setText("65");

        areaChartPass.getData().clear();
        areaChartPassTrim.getData().clear();
        lineChartSeriesOverlay.getData().clear();
        areaChartSeriesAverage.getData().clear();
        areaChartRacetrack.getData().clear();
        areaChartBackAndForth.getData().clear();

        labelSwathFinal.setText("");
        rt01.setText("");
        rt11.setText("");
        rt02.setText("");
        rt12.setText("");
        rt03.setText("");
        rt13.setText("");
        rt04.setText("");
        rt14.setText("");
        rt05.setText("");
        rt15.setText("");
        rt06.setText("");
        rt16.setText("");
        rt07.setText("");
        rt17.setText("");
        rt08.setText("");
        rt18.setText("");
        rt09.setText("");
        rt19.setText("");
        rt010.setText("");
        rt110.setText("");
        rt011.setText("");
        rt111.setText("");
        bf01.setText("");
        bf11.setText("");
        bf02.setText("");
        bf12.setText("");
        bf03.setText("");
        bf13.setText("");
        bf04.setText("");
        bf14.setText("");
        bf05.setText("");
        bf15.setText("");
        bf06.setText("");
        bf16.setText("");
        bf07.setText("");
        bf17.setText("");
        bf08.setText("");
        bf18.setText("");
        bf09.setText("");
        bf19.setText("");
        bf010.setText("");
        bf110.setText("");
        bf011.setText("");
        bf111.setText("");

        clearCardSlate();
    }

    public void updatePass(){
        int pass = (Integer) spinnerPassNum.getValue();
        switch (pass){
            case 1:
                spinnerTrimL.getValueFactory().setValue(patternObject.getTrim1L());
                spinnerTrimR.getValueFactory().setValue(patternObject.getTrim1R());
                sliderVerticalTrim.setValue(patternObject.getTrimVertical1());
                break;
            case 2:
                spinnerTrimL.getValueFactory().setValue(patternObject.getTrim2L());
                spinnerTrimR.getValueFactory().setValue(patternObject.getTrim2R());
                sliderVerticalTrim.setValue(patternObject.getTrimVertical2());
                break;
            case 3:
                spinnerTrimL.getValueFactory().setValue(patternObject.getTrim3L());
                spinnerTrimR.getValueFactory().setValue(patternObject.getTrim3R());
                sliderVerticalTrim.setValue(patternObject.getTrimVertical3());
                break;
            case 4:

                break;
            case 5:

                break;
            case 6:

                break;
        }
        updateVisibilityAnalyzing(false, pass);
        refreshPlots1();
    }

    public void refreshPlots1(){
        int pass = (Integer) spinnerPassNum.getValue();
        areaChartPass.getData().clear();
        areaChartPassTrim.getData().clear();
        if(isScannedPass1){
            if (pass==1) {
                plotRawPass(patternObject.getPass1(),patternObject.getSampleLength());
            }
            patternObject.setPass1Mod(smoothPattern(patternObject.getPass1(),isSmoothSeries));
            patternObject.setPass1Mod(trimAndZeroPattern(patternObject.getPass1Mod(),patternObject.getTrim1L(),patternObject.getTrim1R()));
            if (pass==1) {
                plotTrimmedPass(patternObject.getPass1Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical1());
            }
            patternObject.setPass1Mod(verticallyTrimPattern(patternObject.getPass1Mod(),patternObject.getTrimVertical1()));
            patternObject.setPass1Mod(centroidPattern(patternObject.getPass1Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        if(isScannedPass2){
            if (pass==2) {
                plotRawPass(patternObject.getPass2(),patternObject.getSampleLength());
            }
            patternObject.setPass2Mod(smoothPattern(patternObject.getPass2(),isSmoothSeries));
            patternObject.setPass2Mod(trimAndZeroPattern(patternObject.getPass2Mod(),patternObject.getTrim2L(),patternObject.getTrim2R()));
            if (pass==2) {
                plotTrimmedPass(patternObject.getPass2Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical2());
            }
            patternObject.setPass2Mod(verticallyTrimPattern(patternObject.getPass2Mod(),patternObject.getTrimVertical2()));
            patternObject.setPass2Mod(centroidPattern(patternObject.getPass2Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        if(isScannedPass3){
            if (pass==3) {
                plotRawPass(patternObject.getPass3(),patternObject.getSampleLength());
            }
            patternObject.setPass3Mod(smoothPattern(patternObject.getPass3(),isSmoothSeries));
            patternObject.setPass3Mod(trimAndZeroPattern(patternObject.getPass3Mod(),patternObject.getTrim3L(),patternObject.getTrim3R()));
            if (pass==3) {
                plotTrimmedPass(patternObject.getPass3Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical3());
            }
            patternObject.setPass3Mod(verticallyTrimPattern(patternObject.getPass3Mod(),patternObject.getTrimVertical3()));
            patternObject.setPass3Mod(centroidPattern(patternObject.getPass3Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        if(isScannedPass4){
            if (pass==4) {
                plotRawPass(patternObject.getPass4(),patternObject.getSampleLength());
            }
            patternObject.setPass4Mod(smoothPattern(patternObject.getPass4(),isSmoothSeries));
            patternObject.setPass4Mod(trimAndZeroPattern(patternObject.getPass4Mod(),patternObject.getTrim4L(),patternObject.getTrim4R()));
            if (pass==4) {
                plotTrimmedPass(patternObject.getPass4Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical4());
            }
            patternObject.setPass4Mod(verticallyTrimPattern(patternObject.getPass4Mod(),patternObject.getTrimVertical4()));
            patternObject.setPass4Mod(centroidPattern(patternObject.getPass4Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        if(isScannedPass5){
            if (pass==5) {
                plotRawPass(patternObject.getPass5(),patternObject.getSampleLength());
            }
            patternObject.setPass5Mod(smoothPattern(patternObject.getPass5(),isSmoothSeries));
            patternObject.setPass5Mod(trimAndZeroPattern(patternObject.getPass5Mod(),patternObject.getTrim5L(),patternObject.getTrim5R()));
            if (pass==5) {
                plotTrimmedPass(patternObject.getPass5Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical5());
            }
            patternObject.setPass5Mod(verticallyTrimPattern(patternObject.getPass5Mod(),patternObject.getTrimVertical5()));
            patternObject.setPass5Mod(centroidPattern(patternObject.getPass5Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        if(isScannedPass6){
            if (pass==6) {
                plotRawPass(patternObject.getPass6(),patternObject.getSampleLength());
            }
            patternObject.setPass6Mod(smoothPattern(patternObject.getPass6(),isSmoothSeries));
            patternObject.setPass6Mod(trimAndZeroPattern(patternObject.getPass6Mod(),patternObject.getTrim6L(),patternObject.getTrim6R()));
            if (pass==6) {
                plotTrimmedPass(patternObject.getPass6Mod(), patternObject.getSampleLength(),patternObject.getTrimVertical6());
            }
            patternObject.setPass6Mod(verticallyTrimPattern(patternObject.getPass6Mod(),patternObject.getTrimVertical6()));
            patternObject.setPass6Mod(centroidPattern(patternObject.getPass6Mod(),patternObject.getSampleLength(),isAlignCentroid));
        }
        //Testing flipping
        patternObject.setPatternFlipped(false);
    }

    public void refreshPlots2(){
        updatePassColors();
        if(isEqualizeArea){equalizeArea();}
        plotOverlay();
        plotAverage();
        plotRTCV();
        plotBFCV();
    }

    public void updateTrim(){
        int pass = (Integer) spinnerPassNum.getValue();
        switch (pass){
            case 1:
                patternObject.setTrim1L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim1R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical1(sliderVerticalTrim.getValue());
                break;
            case 2:
                patternObject.setTrim2L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim2R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical2(sliderVerticalTrim.getValue());
                break;
            case 3:
                patternObject.setTrim3L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim3R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical3(sliderVerticalTrim.getValue());
                break;
            case 4:
                patternObject.setTrim4L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim4R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical4(sliderVerticalTrim.getValue());
                break;
            case 5:
                patternObject.setTrim5L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim5R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical5(sliderVerticalTrim.getValue());
                break;
            case 6:
                patternObject.setTrim6L((Integer) spinnerTrimL.getValue());
                patternObject.setTrim6R((Integer) spinnerTrimR.getValue());
                patternObject.setTrimVertical6(sliderVerticalTrim.getValue());
                break;
        }
        refreshPlots1();
        refreshPlots2();
    }

    public void updateVisibilityAnalyzing (Boolean analyzing, Integer pass){
        buttonManualReverse.setDisable(analyzing);
        buttonManualAdvance.setDisable(analyzing);
        spinnerPassNum.setDisable(analyzing);
        spinnerTrimL.setDisable(analyzing);
        spinnerTrimR.setDisable(analyzing);

        if(pass == 1){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass1) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        } else if(pass == 2){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass2) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        } else if(pass == 3){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass3) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        }
        if(pass == 4){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass4) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        } else if(pass == 5){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass5) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        } else if(pass == 6){
            if(!analyzing){
                buttonPassStart.setText("Start");
                if (!isScannedPass6) {
                    buttonPassStart.setDisable(false);
                    //buttonPassStop.setText("Skip");
                } else {
                    buttonPassStop.setText("Repeat");
                    buttonPassStart.setDisable(true);
                }
            } else {
                buttonPassStart.setDisable(true);
                buttonPassStop.setText("Stop");
            }
        }
    }

    public void updatePassColors(){
        if(isScannedPass1){
            toggleButtonPass1.setStyle("-fx-base: #328432;");
        } else {
            toggleButtonPass1.setStyle("-fx-base: gray;");
        }
        if(isScannedPass2){
            toggleButtonPass2.setStyle("-fx-base: #e9967a;");
        } else {
            toggleButtonPass2.setStyle("-fx-base: gray;");
        }
        if(isScannedPass3){
            toggleButtonPass3.setStyle("-fx-base: #ffcc00;");
        } else {
            toggleButtonPass3.setStyle("-fx-base: gray;");
        }
        if(isScannedPass4){
            toggleButtonPass4.setStyle("-fx-base: #328432;");
        } else {
            toggleButtonPass4.setStyle("-fx-base: gray;");
        }
        if(isScannedPass5){
            toggleButtonPass5.setStyle("-fx-base: #e9967a;");
        } else {
            toggleButtonPass5.setStyle("-fx-base: gray;");
        }
        if(isScannedPass6){
            toggleButtonPass6.setStyle("-fx-base: #ffcc00;");
        } else {
            toggleButtonPass6.setStyle("-fx-base: gray;");
        }

        /*
        int[] arr = new int[6];
        if(isScannedPass1){arr[0] = 1;} else {arr[0] = 0;}
        if(isScannedPass2){arr[1] = 1;} else {arr[1] = 0;}
        if(isScannedPass3){arr[2] = 1;} else {arr[2] = 0;}
        if(isScannedPass4){arr[3] = 1;} else {arr[3] = 0;}
        if(isScannedPass5){arr[4] = 1;} else {arr[4] = 0;}
        if(isScannedPass6){arr[5] = 1;} else {arr[5] = 0;}

        int[] add = new int[6];
        add[0] = arr[0];
        if (arr[1]!=0) {
            add[1] = arr[0]+arr[1];
        } else {
            add[1]=0;
        }
        if (arr[2]!=0) {
            add[2] = arr[0]+arr[1]+arr[2];
        } else {
            add[2] = 0;
        }
        if (arr[3]!=0) {
            add[3] = arr[0]+arr[1]+arr[2]+arr[3];
        } else {
            add[3] = 0;
        }
        if (arr[4]!=0) {
            add[4] = arr[0]+arr[1]+arr[2]+arr[3]+arr[4];
        } else {
            add[4] = 0;
        }
        if (arr[5]!=0) {
            add[5] = arr[0]+arr[1]+arr[2]+arr[3]+arr[4]+arr[5];
        } else {
            add[5] = 0;
        }

        String[] col = new String[6];
        for(int i=0; i<6; i++){
            if(add[i]==1 || add[i]==4){
                col[i] = "-fx-base: #328432";
            } else if (add[i]==2 || add[i]==5){
                col[i] = "-fx-base: #e9967a";
            } else if (add[i]==3 || add[i]==6){
                col[i] = "-fx-base: #ffcc00";
            } else {
                col[i] = "-fx-base: gray";
            }
        }
        toggleButtonPass1.setStyle(col[0]);
        toggleButtonPass2.setStyle(col[1]);
        toggleButtonPass3.setStyle(col[2]);
        toggleButtonPass4.setStyle(col[3]);
        toggleButtonPass5.setStyle(col[4]);
        toggleButtonPass6.setStyle(col[5]);
        */
    }

    public void updateActivePasses() {
        textFieldGS1.setDisable(!cBPass1.isSelected());
        textFieldGS2.setDisable(!cBPass2.isSelected());
        textFieldGS3.setDisable(!cBPass3.isSelected());
        textFieldGS4.setDisable(!cBPass4.isSelected());
        textFieldGS5.setDisable(!cBPass5.isSelected());
        textFieldGS6.setDisable(!cBPass6.isSelected());

        textFieldSH1.setDisable(!cBPass1.isSelected());
        textFieldSH2.setDisable(!cBPass2.isSelected());
        textFieldSH3.setDisable(!cBPass3.isSelected());
        textFieldSH4.setDisable(!cBPass4.isSelected());
        textFieldSH5.setDisable(!cBPass5.isSelected());
        textFieldSH6.setDisable(!cBPass6.isSelected());

        textFieldPH1.setDisable(!cBPass1.isSelected());
        textFieldPH2.setDisable(!cBPass2.isSelected());
        textFieldPH3.setDisable(!cBPass3.isSelected());
        textFieldPH4.setDisable(!cBPass4.isSelected());
        textFieldPH5.setDisable(!cBPass5.isSelected());
        textFieldPH6.setDisable(!cBPass6.isSelected());

        textFieldWD1.setDisable(!cBPass1.isSelected());
        textFieldWD2.setDisable(!cBPass2.isSelected());
        textFieldWD3.setDisable(!cBPass3.isSelected());
        textFieldWD4.setDisable(!cBPass4.isSelected());
        textFieldWD5.setDisable(!cBPass5.isSelected());
        textFieldWD6.setDisable(!cBPass6.isSelected());

        textFieldWV1.setDisable(!cBPass1.isSelected());
        textFieldWV2.setDisable(!cBPass2.isSelected());
        textFieldWV3.setDisable(!cBPass3.isSelected());
        textFieldWV4.setDisable(!cBPass4.isSelected());
        textFieldWV5.setDisable(!cBPass5.isSelected());
        textFieldWV6.setDisable(!cBPass6.isSelected());

        activePasses.clear();
        if(cBPass1.isSelected()){activePasses.add(1);}
        if(cBPass2.isSelected()){activePasses.add(2);}
        if(cBPass3.isSelected()){activePasses.add(3);}
        if(cBPass4.isSelected()){activePasses.add(4);}
        if(cBPass5.isSelected()){activePasses.add(5);}
        if(cBPass6.isSelected()){activePasses.add(6);}
    }

    //Manipulate Pattern Data
    private ArrayList<Double> trimAndZeroPattern(ArrayList<Double> passPattern, int trimLeft, int trimRight){
        //ArrayList to output in the end
        ArrayList<Double> newPattern = new ArrayList<>();
        //Cast to double array
        double[] pattern = new double[passPattern.size()];
        for (int i = 0; i < pattern.length; i++) {
            pattern[i] = passPattern.get(i);
        }

        //Find Lowest Point if Horizontal Trim is Applied
        double lowestPoint = 65535.0;
        for (int i = trimLeft; i < pattern.length-trimRight; i++) {
            if(pattern[i]<lowestPoint){
                lowestPoint = pattern[i];
            }
        }
        //Remove area under lowest point and Assign all points outside of trim to lowest point value
        for (int i = 0; i < pattern.length; i++) {
            if(pattern[i]<lowestPoint){
                pattern[i] = 0;
            } else {
                pattern[i] -= lowestPoint;
            }
            if(i<trimLeft || i>pattern.length-trimRight){
                pattern[i] = 0;
            }
            newPattern.add(pattern[i]);
        }

        return newPattern;
    }

    private ArrayList<Double> smoothPatternOld(ArrayList<Double> passPattern, boolean smooth){
        //ArrayList to output in the end
        ArrayList<Double> newPattern = new ArrayList<>();
        //Cast to double array
        double[] pattern = new double[passPattern.size()];
        for (int i = 0; i < pattern.length; i++) {
            pattern[i] = passPattern.get(i);
        }

        //Craete and assign smoothed values to tempPattern
        double[]tempPattern = new double[passPattern.size()];
        for(int i=0; i<pattern.length; i++) {
            //Using quadratic polynomial with 9 points
            if (smooth) {
                if (i == 0) {
                    tempPattern[i] = ((59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3]) - (21 * pattern[i + 4])) / 145;
                } else if (i == 1) {
                    tempPattern[i] = ((54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3]) - (21 * pattern[i + 4])) / 199;
                } else if (i == 2) {
                    tempPattern[i] = ((39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3]) - (21 * pattern[i + 4])) / 238;
                } else if (i == 3) {
                    tempPattern[i] = ((14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3]) - (21 * pattern[i + 4])) / 252;
                } else if (i >= 4 && i < pattern.length - 4) {
                    tempPattern[i] = (-(21 * pattern[i - 4]) + (14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3]) - (21 * pattern[i + 4])) / 231;
                } else if (i == pattern.length - 4) {
                    tempPattern[i] = (-(21 * pattern[i - 4]) + (14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2]) + (14 * pattern[i + 3])) / 252;
                } else if (i == pattern.length - 3) {
                    tempPattern[i] = (-(21 * pattern[i - 4]) + (14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1]) + (39 * pattern[i + 2])) / 238;
                } else if (i == pattern.length - 2) {
                    tempPattern[i] = (-(21 * pattern[i - 4]) + (14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i]) + (54 * pattern[i + 1])) / 199;
                } else if (i == pattern.length - 1) {
                    tempPattern[i] = (-(21 * pattern[i - 4]) + (14 * pattern[i - 3]) + (39 * pattern[i - 2]) + (54 * pattern[i - 1]) + (59 * pattern[i])) / 145;
                } else {
                    tempPattern[i] = 0;
                }
            } else {
                //No Smoothing
                tempPattern[i] = pattern[i];
            }
        }
        //Reassign values from tempPattern to overwrite original pattern (now smoothed)
        double lowpoint = 65535.0;
        for(int i=0; i<pattern.length; i++) {
            pattern[i] = tempPattern[i];
            if(pattern[i]<lowpoint){
                lowpoint = pattern[i];
            }
        }
        //Rezero
        for (int i = 0; i < pattern.length; i++) {
            if(pattern[i]>lowpoint){
                pattern[i] = pattern[i]-lowpoint;
            } else {
                pattern[i] = 0;
            }
            newPattern.add(pattern[i]);
        }

        return newPattern;
    }

    private ArrayList<Double> smoothPattern(ArrayList<Double> passPattern, boolean smooth){
        if (smooth) {
            //Inputs
            double[] pass = new double[passPattern.size()];
            for(int i=0; i<pass.length; i++){
                pass[i] = passPattern.get(i);
            }
            int left = 25;
            int right = 25;
            int order = 4;
            //Internals
            double[] coefficients = savitskyGolayCoefficients(left,right,order);
            int pts = pass.length;
            double[] ret_value = new double[pts];
            int j;
            //Left Side Reflected
            for(j=0; j<left; j++){
                ret_value[j] = reflectiveSavitskyGolayFilter(pass, coefficients, j, left);
            }
            //Everything in the middle
            for(j=left; j<pts-right; j++){
                ret_value[j] = savitskyGolayFilter(pass, coefficients, j, left);
            }
            //Right Side Reflected
            for(j=pts-right; j<pts; j++){
                ret_value[j] = reflectiveSavitskyGolayFilter(pass, coefficients, j, left);
            }
            //Send back to an ArrayList
            ArrayList<Double> passback = new ArrayList<>();
            for(int i=0; i<ret_value.length; i++){
                if (ret_value[i]>0.0) {
                    passback.add(ret_value[i]);
                } else {
                    passback.add(0.0);
                }
            }
            return passback;
        } else {
            return passPattern;
        }

    }

    private ArrayList<Double> verticallyTrimPattern(ArrayList<Double> passPattern, double verticalTrim){
        ArrayList<Double> newPattern = new ArrayList<>();
        for(int i=0; i<passPattern.size(); i++){
            if(passPattern.get(i)>verticalTrim){
                newPattern.add(passPattern.get(i)-verticalTrim);
            } else {
                newPattern.add(0.0);
            }
        }
        return newPattern;
    }

    private ArrayList<Double> centroidPattern(ArrayList<Double> passPattern, double sampleLength, boolean centroidify){
        //ArrayList to output in the end
        ArrayList<Double> newPattern = new ArrayList<>();
        //Cast to double array
        double[] pattern = new double[passPattern.size()];
        for (int i = 0; i < pattern.length; i++) {
            pattern[i] = passPattern.get(i);
        }

        //Craete and assign smoothed values to tempPattern
        double[]tempPattern = new double[passPattern.size()];

        //If centroidify, do it
        if (centroidify) {
            //Find Centroid
            double centroidNumerator1 = 0;
            double centroidDenominator1 = 0;
            double flightlineLength = sampleLength * passPattern.size();
            double startpoint = -(flightlineLength / 2);
            double lowpoint = 65535;
            for (int i = 0; i < pattern.length; i++) {
                centroidNumerator1 += ((pattern[i]) * (startpoint));
                centroidDenominator1 += pattern[i];
                startpoint += sampleLength;
                if(pattern[i]<lowpoint){
                    lowpoint = pattern[i];
                }
            }
            double centroid = centroidNumerator1 / centroidDenominator1;
            //Convert calculated centroid to integer of points to move
            int c1 = (int) (centroid / sampleLength);
            //Write the new centroidified series
            tempPattern = new double[passPattern.size()];
            for (int y = 0; y < pattern.length; y++) {
                if ((y + c1) >= 0 && (y + c1) < pattern.length) {
                    tempPattern[y] = pattern[y + c1];
                } else {
                    tempPattern[y] = lowpoint;
                }
            }
            //Reassign values from tempPattern to overwrite original pattern (now centroidified)
            for (int i = 0; i < pattern.length; i++) {
                pattern[i] = tempPattern[i];
                newPattern.add(pattern[i]);
            }
            return newPattern;
        } else {
            return passPattern;
        }

    }

    private ArrayList<Double> flipPattern(ArrayList<Double> passPattern, boolean flipPattern){
        //ArrayList to output in the end
        ArrayList<Double> newPattern = new ArrayList<>();
        //Cast to double array
        double[] pattern = new double[passPattern.size()];
        for (int i = 0; i < pattern.length; i++) {
            pattern[i] = passPattern.get(i);
        }

        //If flipping, do it
        if (flipPattern) {
            for(int i=pattern.length-1; i>=0; i--){
                newPattern.add(pattern[i]);
            }

            return newPattern;
        } else {
            return passPattern;
        }
    }

    private void equalizeArea(){
        double sum1 = 0;
        if(isScannedPass1){
            for(int i=0; i<patternObject.getPass1Mod().size(); i++){
                sum1+=patternObject.getPass1Mod().get(i);
            }
        }
        double sum2 = 0;
        if(isScannedPass2){
            for(int i=0; i<patternObject.getPass2Mod().size(); i++){
                sum2+=patternObject.getPass2Mod().get(i);
            }
        }
        double sum3 = 0;
        if(isScannedPass3){
            for(int i=0; i<patternObject.getPass3Mod().size(); i++){
                sum3+=patternObject.getPass3Mod().get(i);
            }
        }
        double sum4 = 0;
        if(isScannedPass4){
            for(int i=0; i<patternObject.getPass4Mod().size(); i++){
                sum4+=patternObject.getPass4Mod().get(i);
            }
        }
        double sum5 = 0;
        if(isScannedPass5){
            for(int i=0; i<patternObject.getPass5Mod().size(); i++){
                sum5+=patternObject.getPass5Mod().get(i);
            }
        }
        double sum6 = 0;
        if(isScannedPass6){
            for(int i=0; i<patternObject.getPass6Mod().size(); i++){
                sum6+=patternObject.getPass6Mod().get(i);
            }
        }
        double high = DoubleStream.of(sum1, sum2, sum3, sum4, sum5, sum6).max().getAsDouble();
        if(isScannedPass1){
            for(int i=0; i<patternObject.getPass1Mod().size(); i++){
                patternObject.getPass1Mod().set(i,patternObject.getPass1Mod().get(i)/(sum1/high));
            }
        }
        if(isScannedPass2){
            for(int i=0; i<patternObject.getPass2Mod().size(); i++){
                patternObject.pass2Mod.set(i,patternObject.getPass2Mod().get(i)/(sum2/high));
            }
        }
        if(isScannedPass3){
            for(int i=0; i<patternObject.getPass3Mod().size(); i++){
                patternObject.pass3Mod.set(i,patternObject.getPass3Mod().get(i)/(sum3/high));
            }
        }
        if(isScannedPass4){
            for(int i=0; i<patternObject.getPass4Mod().size(); i++){
                patternObject.pass4Mod.set(i,patternObject.getPass4Mod().get(i)/(sum4/high));
            }
        }
        if(isScannedPass5){
            for(int i=0; i<patternObject.getPass5Mod().size(); i++){
                patternObject.pass5Mod.set(i,patternObject.getPass5Mod().get(i)/(sum5/high));
            }
        }
        if(isScannedPass6){
            for(int i=0; i<patternObject.getPass6Mod().size(); i++){
                patternObject.pass6Mod.set(i,patternObject.getPass6Mod().get(i)/(sum6/high));
            }
        }
    }

    private ArrayList<Double> averagePattern(){
        ArrayList<Double> ave = new ArrayList<>();
        int numPasses = 0;
        //Add Pass1 if applicable
        if(isScannedPass1){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass1Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass1Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass1Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass1Mod().size(); i++){
                    ave.add(patternObject.getPass1Mod().get(i));
                }
            }
            numPasses++;
        }
        //Add Pass2 if applicable
        if(isScannedPass2){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass2Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass2Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass2Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass2Mod().size(); i++){
                    ave.add(patternObject.getPass2Mod().get(i));
                }
            }
            numPasses++;
        }
        //Add Pass3 if applicable
        if(isScannedPass3){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass3Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass3Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass3Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass3Mod().size(); i++){
                    ave.add(patternObject.getPass3Mod().get(i));
                }
            }
            numPasses++;
        }
        //Add Pass4 if applicable
        if(isScannedPass4){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass4Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass4Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass4Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass4Mod().size(); i++){
                    ave.add(patternObject.getPass4Mod().get(i));
                }
            }
            numPasses++;
        }
        //Add Pass5 if applicable
        if(isScannedPass5){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass5Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass5Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass5Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass5Mod().size(); i++){
                    ave.add(patternObject.getPass5Mod().get(i));
                }
            }
            numPasses++;
        }
        //Add Pass6 if applicable
        if(isScannedPass6){
            if(ave.size()>0){
                for(int i=0; i<patternObject.getPass6Mod().size(); i++){
                    if (i<ave.size()) {
                        double oldVal = ave.get(i);
                        ave.set(i,oldVal+patternObject.getPass6Mod().get(i));
                    } else {
                        ave.add(patternObject.getPass6Mod().get(i));
                    }
                }
            } else {
                for(int i=0; i<patternObject.getPass6Mod().size(); i++){
                    ave.add(patternObject.getPass6Mod().get(i));
                }
            }
            numPasses++;
        }
        //Compute Average
        for(int i=0; i<ave.size(); i++){
            ave.set(i,ave.get(i)/numPasses);
        }
        return ave;
    }

    //Put patterns in plot windows
    private void plotRawPass(ArrayList<Double> passPattern, double sampleLength){
        double flightlineLength = sampleLength*passPattern.size();
        double startpoint = -(flightlineLength/2);
        AreaChart.Series pattern = new AreaChart.Series();
        for(int i=0 ; i<passPattern.size(); i++){
            pattern.getData().add(new AreaChart.Data<Number, Number>(startpoint+(i*sampleLength), passPattern.get(i)));
        }
        areaChartPass.getData().clear();
        areaChartPass.getData().addAll(pattern);
    }

    private void plotTrimmedPass(ArrayList<Double> passPattern, double sampleLength, double verticalTrim){
        double flightlineLength = sampleLength*passPattern.size();
        double startpoint = -(flightlineLength/2);
        AreaChart.Series pattern = new AreaChart.Series();
        AreaChart.Series seriesVertTrim = new AreaChart.Series();
        //Find high and low for vertical trim slider
        double highpoint = 0;
        double lowpoint = 65535;
        for(int i=0; i<passPattern.size(); i++){
            if(passPattern.get(i)<lowpoint){
                lowpoint=passPattern.get(i);
            }
            if(passPattern.get(i)>highpoint){
                highpoint=passPattern.get(i);
            }
        }
        sliderVerticalTrim.setMax(highpoint-lowpoint);
        sliderVerticalTrim.setValue(verticalTrim);
        for(int i=0 ; i<passPattern.size(); i++){
            pattern.getData().add(new AreaChart.Data<Number, Number>(startpoint+(i*sampleLength), passPattern.get(i)));
            seriesVertTrim.getData().add(new AreaChart.Data<Number,Number>(startpoint+(i*sampleLength),sliderVerticalTrim.getValue()));
        }
        areaChartPassTrim.getData().clear();
        areaChartPassTrim.getData().addAll(pattern,seriesVertTrim);
    }

    private void plotOverlay(){
        //decide if flip is needed
        boolean flipit = false;
        if (!patternObject.getPatternFlipped() && isFlipPattern) {
            System.out.println("Going to flip pattern");
            flipit = true;
            patternObject.setPatternFlipped(true);
        } else if (!patternObject.getPatternFlipped() && !isFlipPattern) {
            System.out.println("Pattern obj flipped = false, isFlipPattern = false");
        } else if (patternObject.getPatternFlipped()){
            System.out.println("Pattern obj flipped = true");
        } else {
            System.out.println("Some other flipped condition");
        }


        //Clear Overlay Plot and Series
        lineChartSeriesOverlay.getData().clear();
        XYChart.Series seriesOverlay1 = new XYChart.Series();
        XYChart.Series seriesOverlay2 = new XYChart.Series();
        XYChart.Series seriesOverlay3 = new XYChart.Series();
        XYChart.Series seriesOverlay4 = new XYChart.Series();
        XYChart.Series seriesOverlay5 = new XYChart.Series();
        XYChart.Series seriesOverlay6 = new XYChart.Series();
        //Plot each series if togglebuttons are selected.
        double sampleLength = patternObject.getSampleLength();
        if(isScannedPass1){
            double flightlineLength = sampleLength*patternObject.getPass1Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass1Mod(flipPattern(patternObject.getPass1Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass1Mod().size(); i++){
                seriesOverlay1.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass1Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay1);
            seriesOverlay1.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #328432; -fx-stroke-width: 2px;");
        }
        if(isScannedPass2){
            double flightlineLength = sampleLength*patternObject.getPass2Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass2Mod(flipPattern(patternObject.getPass2Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass2Mod().size(); i++){
                seriesOverlay2.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass2Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay2);
            seriesOverlay2.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #e9967a; -fx-stroke-width: 2px;");
        }
        if(isScannedPass3){
            double flightlineLength = sampleLength*patternObject.getPass3Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass3Mod(flipPattern(patternObject.getPass3Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass3Mod().size(); i++){
                seriesOverlay3.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass3Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay3);
            seriesOverlay3.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #ffcc00; -fx-stroke-width: 2px;");
        }
        if(isScannedPass4){
            double flightlineLength = sampleLength*patternObject.getPass4Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass4Mod(flipPattern(patternObject.getPass4Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass4Mod().size(); i++){
                seriesOverlay4.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass4Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay4);
            seriesOverlay4.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #328432; -fx-stroke-width: 1.5px; -fx-stroke-dash-array: 7 5;");
        }
        if(isScannedPass5){
            double flightlineLength = sampleLength*patternObject.getPass5Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass5Mod(flipPattern(patternObject.getPass5Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass5Mod().size(); i++){
                seriesOverlay5.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass5Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay5);
            seriesOverlay5.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #e9967a; -fx-stroke-width: 1.5px; -fx-stroke-dash-array: 7 5;");
        }
        if(isScannedPass6){
            double flightlineLength = sampleLength*patternObject.getPass6Mod().size();
            double startpoint = -(flightlineLength/2);
            //Flip if wanted
            patternObject.setPass6Mod(flipPattern(patternObject.getPass6Mod(), flipit));
            //Make series
            for(int i=0; i<patternObject.getPass6Mod().size(); i++){
                seriesOverlay6.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,patternObject.getPass6Mod().get(i)));
            }
            lineChartSeriesOverlay.getData().add(seriesOverlay6);
            seriesOverlay6.getNode().lookup(".chart-series-line").setStyle("-fx-stroke: #ffcc00; -fx-stroke-width: 1.5px; -fx-stroke-dash-array: 7 5;");

        }
    }

    private void plotAverage(){
        //Clear Average Chart
        areaChartSeriesAverage.getData().clear();
        double sampleLength = patternObject.getSampleLength();
        AreaChart.Series avePattSeries = new AreaChart.Series();
        //Calculate Average Pattern (Y points) and Rezero
        ArrayList<Double> avePatt = trimAndZeroPattern(averagePattern(),0,0);
        //Smooth Average Pattern
        avePatt = smoothPattern(avePatt, isSmoothAverage);
        //Centroid Average Pattern
        avePatt = centroidPattern(avePatt,sampleLength,isAlignCentroid);
        //Do one final zeroing
        avePatt = trimAndZeroPattern(avePatt, 0, 0);
        //Copy the final average pattern to the pattern object
        patternObject.setPassAverage(avePatt);
        //Get X points
        double flightlineLength = sampleLength*avePatt.size();
        double startpoint = -(flightlineLength/2);
        //Create Series of X,Y
        for(int i=0; i<avePatt.size(); i++){
            avePattSeries.getData().add(new XYChart.Data<> (startpoint+i*sampleLength,avePatt.get(i)));
        }
        //Put Series into Average Plot
        areaChartSeriesAverage.getData().add(avePattSeries);
        //Find Crosshair
        if (isSwathBox) {
            double swath = Double.parseDouble(labelSwathFinal.getText());
            double crosshairLeft = (flightlineLength/2)-(swath/2);
            double crosshairRight = (flightlineLength/2)+(swath/2);
            double crosshairSum = 0;
            int crosshairPts = 0;
            for(int i=0; i<avePatt.size(); i++){
                if(i*sampleLength>crosshairLeft && i*sampleLength<crosshairRight){
                    crosshairSum += avePatt.get(i);
                    crosshairPts ++;
                }
            }
            //Create Series for Crosshair
            AreaChart.Series seriesCrossHair = new AreaChart.Series();
            seriesCrossHair.getData().add(new XYChart.Data<Number, Number> (crosshairLeft-(flightlineLength/2), 0));
            seriesCrossHair.getData().add(new XYChart.Data<Number, Number> (crosshairLeft-(flightlineLength/2), (crosshairSum/crosshairPts)/2));
            seriesCrossHair.getData().add(new XYChart.Data<Number, Number> (crosshairRight-(flightlineLength/2), (crosshairSum/crosshairPts)/2));
            seriesCrossHair.getData().add(new XYChart.Data<Number, Number> (crosshairRight-(flightlineLength/2), 0));
            //Put Series onto Average Plot
            areaChartSeriesAverage.getData().add(seriesCrossHair);
            numberAxisSeriesAveragey.setLowerBound(0.0);
        }
    }

    private void plotRTCV(){
        double sw = Double.parseDouble(labelSwathFinal.getText());
        double sl = patternObject.getSampleLength();
        double fl = sl*patternObject.getPassAverage().size();
        double startPoint = -(fl/2);
        int swPts = (int) Math.round(sw/sl);
        double boundL = fl-((fl/2)+(sw/2));
        double boundR = fl-((fl/2)-(sw/2));
        double sumPoint;
        double sumPointTotal = 0.0;
        ArrayList<Double> sumPoints = new ArrayList<>();
        ArrayList<Double> locPoints = new ArrayList<>();
        int numPoints = 0;
        //Initialize each series
        areaChartRacetrack.getData().clear();
        AreaChart.Series seriesC = new StackedAreaChart.Series();
        AreaChart.Series seriesL1 = new StackedAreaChart.Series();
        AreaChart.Series seriesL2 = new StackedAreaChart.Series();
        AreaChart.Series seriesR1 = new StackedAreaChart.Series();
        AreaChart.Series seriesR2 = new StackedAreaChart.Series();
        AreaChart.Series seriesElim = new StackedAreaChart.Series();
        AreaChart.Series seriesCrosshair = new StackedAreaChart.Series();
        //Get each point contribution
        for(int i=0; i<patternObject.getPassAverage().size(); i++){
            if (i*sl>=boundL && i*sl<=boundR) {
                //Add orignal average
                seriesC.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i)));
                sumPoint = patternObject.getPassAverage().get(i);
                //Add 1st left
                if(i-swPts>=0 && cV3){
                    seriesL1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i-swPts)));
                    sumPoint += patternObject.getPassAverage().get(i-swPts);
                } else {
                    seriesL1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 2nd Left
                if(i-(2*swPts)>=0 && cV5){
                    seriesL2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i-2*swPts)));
                    sumPoint += patternObject.getPassAverage().get(i-2*swPts);
                } else {
                    seriesL2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 1st Right
                if(i+swPts<patternObject.getPassAverage().size() && cV3){
                    seriesR1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i+swPts)));
                    sumPoint += patternObject.getPassAverage().get(i+swPts);
                } else {
                    seriesR1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 2nd Right
                if(i+(2*swPts)<patternObject.getPassAverage().size() && cV5){
                    seriesR2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i+2*swPts)));
                    sumPoint += patternObject.getPassAverage().get(i+2*swPts);
                } else {
                    seriesR2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                numPoints++;
                sumPoints.add(sumPoint);
                locPoints.add(startPoint+(i*sl));
                sumPointTotal += sumPoint;
                seriesElim.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),-sumPoint));
            }
        }
        //Draw Crosshair
        seriesCrosshair.getData().add(new XYChart.Data<Number,Number>(locPoints.get(0),sumPointTotal/numPoints));
        seriesCrosshair.getData().add(new XYChart.Data<Number,Number>(locPoints.get(locPoints.size()-1),sumPointTotal/numPoints));
        areaChartRacetrack.getData().addAll(seriesC,seriesL1,seriesR1,seriesL2,seriesR2, seriesElim, seriesCrosshair);


        //Get RT Values for Table
        ArrayList<Label> rtSWs = new ArrayList<>();
        ArrayList<Label> rtCVs = new ArrayList<>();
        rtSWs.add(rt01);
        rtCVs.add(rt11);
        rtSWs.add(rt02);
        rtCVs.add(rt12);
        rtSWs.add(rt03);
        rtCVs.add(rt13);
        rtSWs.add(rt04);
        rtCVs.add(rt14);
        rtSWs.add(rt05);
        rtCVs.add(rt15);
        rtSWs.add(rt06);
        rtCVs.add(rt16);
        rtSWs.add(rt07);
        rtCVs.add(rt17);
        rtSWs.add(rt08);
        rtCVs.add(rt18);
        rtSWs.add(rt09);
        rtCVs.add(rt19);
        rtSWs.add(rt010);
        rtCVs.add(rt110);
        rtSWs.add(rt011);
        rtCVs.add(rt111);

        if(!METRIC) {
            for(int i=0; i<rtCVs.size(); i++){
                int sW = (int) Math.round(sw+(-10+(i*2)));
                rtSWs.get(i).setText(Integer.toString(sW));
                if (calcRTCV(sW)>0) {
                    rtCVs.get(i).setText(Integer.toString(calcRTCV(sW))+" %");
                } else {
                    rtCVs.get(i).setText("N/A");
                }
            }
        } else {
            for(int i=0; i<rtCVs.size(); i++){
                double sW = sw+(-2.5+(i*0.5));
                rtSWs.get(i).setText(String.format("%.1f",sW));
                if (calcRTCV(sW)>0) {
                    rtCVs.get(i).setText(Integer.toString(calcRTCV(sW))+" %");
                } else {
                    rtCVs.get(i).setText("N/A");
                }
            }
        }
    }

    private void plotBFCV(){
        double sw = Double.parseDouble(labelSwathFinal.getText());
        double sl = patternObject.getSampleLength();
        double fl = sl*patternObject.getPassAverage().size();
        double startPoint = -(fl/2);
        int swPts = (int) Math.round(sw/sl);
        int startPt = (int) Math.round((patternObject.getPassAverage().size()/2)-(swPts/2));
        double boundL = (fl/2)-(sw/2);
        double boundR = (fl/2)+(sw/2);
        double sumPoint;
        double sumPointTotal = 0.0;
        ArrayList<Double> sumPoints = new ArrayList<>();
        ArrayList<Double> locPoints = new ArrayList<>();
        int numPoints = 0;
        //Initialize each series
        areaChartBackAndForth.getData().clear();
        AreaChart.Series seriesC = new StackedAreaChart.Series();
        AreaChart.Series seriesL1 = new StackedAreaChart.Series();
        AreaChart.Series seriesL2 = new StackedAreaChart.Series();
        AreaChart.Series seriesR1 = new StackedAreaChart.Series();
        AreaChart.Series seriesR2 = new StackedAreaChart.Series();
        AreaChart.Series seriesElim = new StackedAreaChart.Series();
        AreaChart.Series seriesCrosshair = new StackedAreaChart.Series();
        int z=0;
        for(int i=0; i<patternObject.getPassAverage().size(); i++){
            if (i*sl>boundL && i*sl<boundR) {
                //Add orignal average
                seriesC.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(i)));
                sumPoint = patternObject.getPassAverage().get(i);
                //Add 1st left
                if(startPt-z>=0 && cV3){
                    seriesL1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(startPt-z)));
                    sumPoint += patternObject.getPassAverage().get(startPt-z);
                } else {
                    seriesL1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 2nd Left
                if(startPt-swPts-z>=0 && cV5){
                    seriesL2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(startPt-swPts-z)));
                    sumPoint += patternObject.getPassAverage().get(startPt-swPts-z);
                } else {
                    seriesL2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 1st Right
                if(startPt+2*swPts-z<patternObject.getPassAverage().size() && cV3){
                    seriesR1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(startPt+swPts+swPts-z)));
                    sumPoint += patternObject.getPassAverage().get(startPt+swPts+swPts-z);
                } else {
                    seriesR1.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                //Add 2nd Right
                if(startPt+3*swPts-z<patternObject.getPassAverage().size() && cV5){
                    seriesR2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),patternObject.getPassAverage().get(startPt+3*swPts-z)));
                    sumPoint = patternObject.getPassAverage().get(startPt+3*swPts-z);
                } else {
                    seriesR2.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),0));
                }
                numPoints++;
                sumPoints.add(sumPoint);
                locPoints.add(startPoint+(i*sl));
                sumPointTotal += sumPoint;
                seriesElim.getData().add(new XYChart.Data<Number,Number>(startPoint+(i*sl),-sumPoint));
                z++;
            }
        }
        //Draw Crosshair
        seriesCrosshair.getData().add(new XYChart.Data<Number,Number>(locPoints.get(0),sumPointTotal/numPoints));
        seriesCrosshair.getData().add(new XYChart.Data<Number,Number>(locPoints.get(locPoints.size()-1),sumPointTotal/numPoints));
        areaChartBackAndForth.getData().addAll(seriesC,seriesL1,seriesR1,seriesL2,seriesR2,seriesElim,seriesCrosshair);

        ArrayList<Label> bfSWs = new ArrayList<>();
        ArrayList<Label> bfCVs = new ArrayList<>();
        bfSWs.add(bf01);
        bfCVs.add(bf11);
        bfSWs.add(bf02);
        bfCVs.add(bf12);
        bfSWs.add(bf03);
        bfCVs.add(bf13);
        bfSWs.add(bf04);
        bfCVs.add(bf14);
        bfSWs.add(bf05);
        bfCVs.add(bf15);
        bfSWs.add(bf06);
        bfCVs.add(bf16);
        bfSWs.add(bf07);
        bfCVs.add(bf17);
        bfSWs.add(bf08);
        bfCVs.add(bf18);
        bfSWs.add(bf09);
        bfCVs.add(bf19);
        bfSWs.add(bf010);
        bfCVs.add(bf110);
        bfSWs.add(bf011);
        bfCVs.add(bf111);

        if(!METRIC) {
            for(int i=0; i<bfCVs.size(); i++){
                int sW = (int) Math.round(sw+(-10+(i*2)));
                bfSWs.get(i).setText(Integer.toString(sW));
                if (calcBFCV(sW)>0) {
                    bfCVs.get(i).setText(Integer.toString(calcBFCV(sW))+" %");
                } else {
                    bfCVs.get(i).setText("N/A");
                }
            }
        } else {
            for(int i=0; i<bfCVs.size(); i++){
                double sW = sw+(-2.5+(i*0.5));
                bfSWs.get(i).setText(String.format("%.1f",sW));
                if (calcBFCV(sW)>0) {
                    bfCVs.get(i).setText(Integer.toString(calcBFCV(sW))+" %");
                } else {
                    bfCVs.get(i).setText("N/A");
                }
            }
        }
    }

    //Calculating CV
    private int calcRTCV(double swathWidth){
        int rtCV = 0;
        double sl = patternObject.getSampleLength();
        double fl = sl*patternObject.getPassAverage().size();
        int swPts = (int) Math.round(swathWidth/sl);
        double boundL = (fl/2)-(swathWidth/2);
        double boundR = (fl/2)+(swathWidth/2);
        ArrayList<Double> sumPatt = new ArrayList<>();
        for(int i=0; i<patternObject.getPassAverage().size(); i++){
            if (i*sl>boundL && i*sl<boundR) {
                //Add orignal average
                double toAdd = patternObject.getPassAverage().get(i);
                //Add 1st left
                if(i-swPts>=0 && cV3){
                    toAdd += patternObject.getPassAverage().get(i-swPts);
                }
                //Add 2nd Left
                if(i-(2*swPts)>=0 && cV5){
                    toAdd += patternObject.getPassAverage().get(i-2*swPts);
                }
                //Add 1st Right
                if(i+swPts<patternObject.getPassAverage().size() && cV3){
                    toAdd += patternObject.getPassAverage().get(i+swPts);
                }
                //Add 2nd Right
                if(i+(2*swPts)<patternObject.getPassAverage().size() && cV5){
                    toAdd += patternObject.getPassAverage().get(i+swPts);
                }
                sumPatt.add(toAdd);
            }
        }
        //Calculate CV from Pattern above
        double sum = 0;
        for(int i = 0; i<sumPatt.size(); i++){
            sum += sumPatt.get(i);
        }
        double xBar = sum/(sumPatt.size());
        double sum1 = 0;
        for(int i = 0; i<sumPatt.size(); i++){
            sum1 += Math.pow((sumPatt.get(i) - xBar), 2);
        }
        double stdDev = Math.sqrt(sum1/sumPatt.size());
        rtCV = (int) Math.round((stdDev*100)/xBar);
        return rtCV;
    }

    private int calcBFCV(double swathWidth){
        int bfCV = 0;
        double sl = patternObject.getSampleLength();
        double fl = sl*patternObject.getPassAverage().size();
        int swPts = (int) Math.round(swathWidth/sl);
        int startPt = (int) Math.round((patternObject.getPassAverage().size()-swPts)/2);
        double boundL = (fl/2)-(swathWidth/2);
        double boundR = (fl/2)+(swathWidth/2);
        ArrayList<Double> sumPatt = new ArrayList<>();
        int z=0;
        for(int i=0; i<patternObject.getPassAverage().size(); i++){
            if (i*sl>boundL && i*sl<boundR) {
                //Add orignal average
                double toAdd = patternObject.getPassAverage().get(i);
                //Add 1st left
                if(startPt-z>=0 && cV3){
                    toAdd += patternObject.getPassAverage().get(startPt-z);
                }
                //Add 2nd Left
                if(startPt-swPts-z>=0 && cV5){
                    toAdd += patternObject.getPassAverage().get(startPt-swPts-z);
                }
                //Add 1st Right
                if(startPt+2*swPts-z<patternObject.getPassAverage().size() && cV3){
                    toAdd += patternObject.getPassAverage().get(startPt+2*swPts-z);
                }
                //Add 2nd Right
                if(startPt+3*swPts-z<patternObject.getPassAverage().size() && cV5){
                    toAdd += patternObject.getPassAverage().get(startPt+3*swPts-z);
                }
                sumPatt.add(toAdd);
                z++;
            }
        }
        //Calculate CV from Pattern above
        double sum = 0;
        for(int i = 0; i<sumPatt.size(); i++){
            sum += sumPatt.get(i);
        }
        double xBar = sum/(sumPatt.size());
        double sum1 = 0;
        for(int i = 0; i<sumPatt.size(); i++){
            sum1 += Math.pow((sumPatt.get(i) - xBar), 2);
        }
        double stdDev = Math.sqrt(sum1/sumPatt.size());
        bfCV = (int) Math.round((stdDev*100)/xBar);
        return bfCV;
    }

    //OLD STUFF

    private void resetHorizontalAxes(double flightlineLength){
        numberAxisPassX.setLowerBound(-flightlineLength/2);
        numberAxisPassX.setUpperBound(flightlineLength/2);
        numberAxisPassTrimX.setLowerBound(-flightlineLength/2);
        numberAxisPassTrimX.setUpperBound(flightlineLength/2);
        numberAxisSeriesOverlayx.setLowerBound(-flightlineLength/2);
        numberAxisSeriesOverlayx.setUpperBound(flightlineLength/2);
        numberAxisSeriesAveragex.setLowerBound(-flightlineLength/2);
        numberAxisSeriesAveragex.setUpperBound(flightlineLength/2);
        numberAxisBackAndForthx.setLowerBound(-flightlineLength/2);
        numberAxisBackAndForthx.setUpperBound(flightlineLength/2);
        numberAxisRacetrackx.setLowerBound(-flightlineLength/2);
        numberAxisRacetrackx.setUpperBound(flightlineLength/2);

    }

    //AccuStain Components and Methods ********************************************************************************

    //NewStuff
    private void connectScanner(int deviceIndex){

        //ToDo AccuStain=uncomment
        scanDevice = devices.get(deviceIndex);
        scanDevice.setFileName("Preview.tif");

        labelScannerName.setText(scanDevice.toString());
        if(scanDevice instanceof eu.gnome.morena.Scanner){
            scanner = (eu.gnome.morena.Scanner) scanDevice;
            scanner.setMode(eu.gnome.morena.Scanner.RGB_8);
            System.out.println("instance of morena scanner");
        }
        //
    }

    public void clickButtonPreview(){
        //Check if already ran cards
        boolean abort = false;
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            fis.close();

            if(wb.getSheet("Card Data")!=null){
                Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                alert.setTitle("Confirmation Required");
                alert.setHeaderText("Re-Scan Cards?");
                alert.setContentText("All previous data will be lost.");
                Optional<ButtonType> result = alert.showAndWait();
                if (result.get() == ButtonType.OK){
                    clearCardSlate();
                    cardObject = new CardObject(DEFAULT_THRESHOLD);
                    //imp=null;
                    for(int i=wb.getNumberOfSheets()-1; i>=4; i--){
                        wb.removeSheetAt(i);
                    }
                    FileOutputStream fos = new FileOutputStream(currentDataFile);
                    wb.write(fos);
                    fos.close();
                } else {
                    abort = true;
                }
            }


        } catch (Exception e) {e.printStackTrace();}

        //If we haven't scanned cards yet, let's find a scanner and give it a go
        if (!abort) {
            try {
                //ToDo AccuStain=uncomment
                if(scanDevice==null && Manager.getInstance().listDevices().size()>0){
                    Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
                    alert.setTitle("Confirm Scanner Selection");
                    alert.setHeaderText("Use "+ Manager.getInstance().listDevices().get(scannerNum) + "?");
                    alert.setContentText("To confirm: click OK \nOtherwise, click CANCEL and go to Settings->Scanner to select another device");
                    alert.showAndWait().ifPresent(response -> {
                        if (response == ButtonType.OK) {
                            initialValidateScannersList();
                            ((CheckMenuItem) menuScanners.getItems().get(scannerNum)).setSelected(true);
                            progressBarScan.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
                            connectScanner(scannerNum);
                            fullScan = false;
                            scanIt(PREVIEW_DPI);
                        }
                    });
                } else if (scanDevice!=null){
                    progressBarScan.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
                    fullScan = false;
                    scanIt(PREVIEW_DPI);
                } else {
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setHeaderText("No Scanner Found");
                    alert.setContentText("Check connection to Scanner and try again");
                    alert.showAndWait();
                }

            } catch (Exception e) {
                Alert a = new Alert(Alert.AlertType.ERROR, "Unable to find Scanner.");
                a.showAndWait();
            }
        }
    }

    public void clickButtonFullScan(){
        fullScan = true;

        if(cardObject.getRm().getCount() == checkSelectedCards()){
            //imp.getWindow().setVisible(false);
            cardObject.getImp().getWindow().setVisible(false);
            progressBarScan.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
            scanIt(FULLSCAN_DPI);
            //buttonPreview.setDisable(true);
            //buttonFullScan.setDisable(true);
        } else {
            cardObject.getImp().getWindow().setVisible(false);
            Alert alert = new Alert(Alert.AlertType.ERROR, "Nuumber of cards detected = "+cardObject.getRm().getCount()+"\n"+"Number of cards searching for = "+checkSelectedCards()+"\n"+"Modify options and re-initiate Preview Scan");
            alert.showAndWait();
        }
    }

    public void clickCardToggle(){
//        System.out.println("clickCardToggleBefore "+togL32.isSelected()+" "+togL24.isSelected()+" "+togL16.isSelected());
//        System.out.println("clickCardToggleBefore1 "+cardObject.getCardList().get(0)+" "+cardObject.getCardList().get(1)+" "+cardObject.getCardList().get(2));
//        System.out.println("clickCardToggleBefore2 "+cardObject.getUsingCardList().get(0)+" "+cardObject.getUsingCardList().get(1)+" "+cardObject.getUsingCardList().get(2));
        if (cardObject!=null && cardObject.getUsingCardList().size()>0) {
            if (cardObject.getCardList().get(0)) {
//                System.out.println("TestThis");
                cardObject.getUsingCardList().set(0,togL32.isSelected());
            } else {
                togL32.setSelected(false);
            }
            if (cardObject.getCardList().get(1)) {
                cardObject.getUsingCardList().set(1,togL24.isSelected());
            } else {
                togL24.setSelected(false);
            }
            if (cardObject.getCardList().get(2)) {
                cardObject.getUsingCardList().set(2,togL16.isSelected());
            } else {
                togL16.setSelected(false);
            }
            if (cardObject.getCardList().get(3)) {
                cardObject.getUsingCardList().set(3,togL8.isSelected());
            } else {
                togL8.setSelected(false);
            }
            if (cardObject.getCardList().get(4)) {
                cardObject.getUsingCardList().set(4,tog0.isSelected());
            } else {
                tog0.setSelected(false);
            }
            if (cardObject.getCardList().get(5)) {
                cardObject.getUsingCardList().set(5,togR8.isSelected());
            } else {
                togR8.setSelected(false);
            }
            if (cardObject.getCardList().get(6)) {
                cardObject.getUsingCardList().set(6,togR16.isSelected());
            } else {
                togR16.setSelected(false);
            }
            if (cardObject.getCardList().get(7)) {
                cardObject.getUsingCardList().set(7,togR24.isSelected());
            } else {
                togR24.setSelected(false);
            }
            if (cardObject.getCardList().get(8)) {
                cardObject.getUsingCardList().set(8,togR32.isSelected());
            } else {
                togR32.setSelected(false);
            }
//            System.out.println("clickCardToggleAfter "+togL32.isSelected()+" "+togL24.isSelected()+" "+togL16.isSelected());
//            System.out.println("clickCardToggleAfter2 "+cardObject.getUsingCardList().get(0)+" "+cardObject.getUsingCardList().get(1)+" "+cardObject.getUsingCardList().get(2));
            updateViewUsingCards();
            showComposites();
            updateDataFileUsingCards();
        }
    }

    private void updateViewUsingCards(){
        if (cardObject.getUsingCardList().get(0)) {
            vBoxCard0.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard0T.setVisible(true);
        } else {
            vBoxCard0.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard0T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(1)) {
            vBoxCard1.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard1T.setVisible(true);
        } else {
            vBoxCard1.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard1T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(2)) {
            vBoxCard2.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard2T.setVisible(true);
        } else {
            vBoxCard2.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard2T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(3)) {
            vBoxCard3.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard3T.setVisible(true);
        } else {
            vBoxCard3.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard3T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(4)) {
            vBoxCard4.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard4T.setVisible(true);
        } else {
            vBoxCard4.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard4T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(5)) {
            vBoxCard5.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard5T.setVisible(true);
        } else {
            vBoxCard5.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard5T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(6)) {
            vBoxCard6.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard6T.setVisible(true);
        } else {
            vBoxCard6.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard6T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(7)) {
            vBoxCard7.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard7T.setVisible(true);
        } else {
            vBoxCard7.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard7T.setVisible(false);
        }
        if (cardObject.getUsingCardList().get(8)) {
            vBoxCard8.setStyle("-fx-background-color:  #7fb27f;");
            hBoxCard8T.setVisible(true);
        } else {
            vBoxCard8.setStyle("-fx-background-color: #f1bdac;");
            hBoxCard8T.setVisible(false);
        }
    }

    private void updateDataFileUsingCards(){
        try{
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            fis.close();

            Sheet sh = wb.getSheet("Card Data");

            for(int i=0; i<cardObject.getCardList().size(); i++){
                sh.getRow(2).getCell(4+i).setCellValue(cardObject.getUsingCardList().get(i));
            }

            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    private int checkSelectedCards(){
        int selectedCardCount = 0;
        cardObject.getCardNameList().clear();
        int s = CARD_SPACING;
        if(togL32.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("L-"+Integer.toString(s*4));
        }
        if(togL24.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("L-"+Integer.toString(s*3));
        }
        if(togL16.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("L-"+Integer.toString(s*2));
        }
        if(togL8.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("L-"+Integer.toString(s));
        }
        if(tog0.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("CENTER");
        }
        if(togR8.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("R-"+Integer.toString(s));
        }
        if(togR16.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("R-"+Integer.toString(s*2));
        }
        if(togR24.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("R-"+Integer.toString(s*3));
        }
        if(togR32.isSelected()){
            selectedCardCount++;
            cardObject.getCardNameList().add("R-"+Integer.toString(s*4));
        }
        return selectedCardCount;
    }

    private void setCardSpacing(){
        int s = CARD_SPACING;
        togL32.setText("L"+Integer.toString(s*4));
        togL24.setText("L"+Integer.toString(s*3));
        togL16.setText("L"+Integer.toString(s*2));
        togL8.setText("L"+Integer.toString(s));
        tog0.setText("00");
        togR8.setText("R"+Integer.toString(s));
        togR16.setText("R"+Integer.toString(s*2));
        togR24.setText("R"+Integer.toString(s*3));
        togR32.setText("R"+Integer.toString(s*4));
    }

    private void clearCardSlate(){

        ableToReRun = false;

        imageViewCard0.setImage(null);
        imageViewCard1.setImage(null);
        imageViewCard2.setImage(null);
        imageViewCard3.setImage(null);
        imageViewCard4.setImage(null);
        imageViewCard5.setImage(null);
        imageViewCard6.setImage(null);
        imageViewCard7.setImage(null);
        imageViewCard8.setImage(null);

        imageViewCard0T.setImage(null);
        imageViewCard1T.setImage(null);
        imageViewCard2T.setImage(null);
        imageViewCard3T.setImage(null);
        imageViewCard4T.setImage(null);
        imageViewCard5T.setImage(null);
        imageViewCard6T.setImage(null);
        imageViewCard7T.setImage(null);
        imageViewCard8T.setImage(null);

        labelCard0.setText("");
        labelCard1.setText("");
        labelCard2.setText("");
        labelCard3.setText("");
        labelCard4.setText("");
        labelCard5.setText("");
        labelCard6.setText("");
        labelCard7.setText("");
        labelCard8.setText("");

        labelCard0T.setText("");
        labelCard1T.setText("");
        labelCard2T.setText("");
        labelCard3T.setText("");
        labelCard4T.setText("");
        labelCard5T.setText("");
        labelCard6T.setText("");
        labelCard7T.setText("");
        labelCard8T.setText("");

        labelVMD0.setText("");
        labelVMD1.setText("");
        labelVMD2.setText("");
        labelVMD3.setText("");
        labelVMD4.setText("");
        labelVMD5.setText("");
        labelVMD6.setText("");
        labelVMD7.setText("");
        labelVMD8.setText("");

        labelDrops0.setText("");
        labelDrops1.setText("");
        labelDrops2.setText("");
        labelDrops3.setText("");
        labelDrops4.setText("");
        labelDrops5.setText("");
        labelDrops6.setText("");
        labelDrops7.setText("");
        labelDrops8.setText("");

        labelCoverage0.setText("");
        labelCoverage1.setText("");
        labelCoverage2.setText("");
        labelCoverage3.setText("");
        labelCoverage4.setText("");
        labelCoverage5.setText("");
        labelCoverage6.setText("");
        labelCoverage7.setText("");
        labelCoverage8.setText("");

        labelCCov.setText("");
        labelCVMD.setText("");
        labelCDV01.setText("");
        labelCDV09.setText("");
        labelCRS.setText("");
        labelCDSC.setText("");

        vBoxCard0.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard1.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard2.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard3.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard4.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard5.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard6.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard7.setStyle("-fx-background-color: #f4f4f4;");
        vBoxCard8.setStyle("-fx-background-color: #f4f4f4;");

        hBoxCard0T.setVisible(false);
        hBoxCard1T.setVisible(false);
        hBoxCard2T.setVisible(false);
        hBoxCard3T.setVisible(false);
        hBoxCard4T.setVisible(false);
        hBoxCard5T.setVisible(false);
        hBoxCard6T.setVisible(false);
        hBoxCard7T.setVisible(false);
        hBoxCard8T.setVisible(false);

        areaChartCoverage.getData().clear();
        barChartDropletVolume.getData().clear();
    }

    private void scanIt(int dpi){
        //ToDo AccuStain=uncomment
        try {
            if(cardObject!=null && cardObject.getImp()!=null && cardObject.getImp().isVisible()){
                cardObject.getImp().getWindow().setVisible(false);
            }
            scanner.setResolution(dpi);
            scanner.startTransfer(this);
        } catch (Exception e) {
            e.printStackTrace();
            buttonPreview.setDisable(false);
            buttonFullScan.setDisable(false);
        }
    }

    private void showPreviewScan(File previewScanFile){
        cardObject = new CardObject(DEFAULT_THRESHOLD);
        Opener opener = new Opener();
        cardObject.setImp(opener.openImage(previewScanFile.getPath()));
        cardObject.setImp(generateROIS(cardObject.getImp()));
        cardObject.getImp().show();
        cardObject.getImp().getWindow().toFront();

    }

    private ImagePlus generateROIS(ImagePlus imagePlus){
        ImagePlus imp0 = imagePlus.duplicate();
        //Find Edges
        ImageProcessor ip = imp0.getProcessor();
        ip.findEdges();
        //Convert to 8gray
        ImageConverter ic = new ImageConverter(imp0);
        ic.convertToGray8();
        imp0.updateImage();
        ip = imp0.getProcessor();
        ip.autoThreshold();
        //Test View
        cardObject.setRm(RoiManager.getRoiManager());
        //ParticleAnalyzer to Locate Card Borders
        ResultsTable rt = ResultsTable.getResultsTable();
        ParticleAnalyzer analyzer = new ParticleAnalyzer(ParticleAnalyzer.ADD_TO_MANAGER+ParticleAnalyzer.CLEAR_WORKSHEET+ParticleAnalyzer.EXCLUDE_EDGE_PARTICLES,
                1, rt,
                10000.0D,
                1.7976931348623157E308D,
                0.0D,
                1.0D);
        analyzer.analyze(imp0);
        imp0.updateImage();

        Roi[] rois = cardObject.getRm().getRoisAsArray();
        //Roi[] cropRois = new Roi[rois.length];

        //Test renaming rois
        checkSelectedCards();

        cardObject.setOverlay(new Overlay());
        for(int i=0; i<rois.length; i++){
            cardObject.getRm().select(i);
            cardObject.getRm().runCommand("Delete");

            java.awt.Rectangle rectOG = rois[i].getBounds();
            int xcoord = rectOG.x + (int) (rectOG.width- CROP_SCALE *rectOG.width)/2;
            int ycoord = rectOG.y + (int) (rectOG.height- CROP_SCALE *rectOG.height)/2;
            int width = (int) (rectOG.width* CROP_SCALE);
            int height = (int) (rectOG.height* CROP_SCALE);
            java.awt.Rectangle rectNEW = new java.awt.Rectangle(xcoord,ycoord,width,height);
            rois[i] = new Roi(rectNEW);
            cardObject.getOverlay().add(rois[i]);

            cardObject.getRm().addRoi(cardObject.getOverlay().get(i));

        }
        cardObject.getRm().runCommand("Select All");
        cardObject.getRm().runCommand("Delete");
        for(int i=0; i<rois.length; i++){
            if (i<cardObject.getCardNameList().size()) {
                //If/Else to rename Center to C
                if (cardObject.getCardNameList().get(i).equalsIgnoreCase("Center")) {
                    cardObject.getOverlay().get(i).setName("C");
                    cardObject.getRm().addRoi(cardObject.getOverlay().get(i));
                    cardObject.getRm().getRoi(i).setName("C");
                } else {
                    cardObject.getOverlay().get(i).setName(cardObject.getCardNameList().get(i));
                    cardObject.getRm().addRoi(cardObject.getOverlay().get(i));
                    cardObject.getRm().getRoi(i).setName(cardObject.getCardNameList().get(i));
                }
            } else {
                cardObject.getOverlay().get(i).setName("?");
                cardObject.getRm().addRoi(cardObject.getOverlay().get(i));
                cardObject.getRm().getRoi(i).setName("?");
            }

        }

        cardObject.getRm().setVisible(false);

        cardObject.getOverlay().drawNames(true);
        cardObject.getOverlay().drawLabels(true);
        cardObject.getOverlay().setLabelColor(Color.BLACK);
        cardObject.getOverlay().setStrokeColor(Color.BLACK);
        Font font = new Font("Verdana", Font.BOLD, 12);
        cardObject.getOverlay().setLabelFont(font);
        imagePlus.setOverlay(cardObject.getOverlay());
        return imagePlus;
    }

    private void assignScannedCards(){
        cardObject.getCardList().add(togL32.isSelected());
        cardObject.getCardList().add(togL24.isSelected());
        cardObject.getCardList().add(togL16.isSelected());
        cardObject.getCardList().add(togL8.isSelected());
        cardObject.getCardList().add(tog0.isSelected());
        cardObject.getCardList().add(togR8.isSelected());
        cardObject.getCardList().add(togR16.isSelected());
        cardObject.getCardList().add(togR24.isSelected());
        cardObject.getCardList().add(togR32.isSelected());

        cardObject.getUsingCardList().addAll(cardObject.getCardList());
        updateViewUsingCards();
    }

    private javafx.scene.image.Image originalSnapshot(int indexOfROI){
        return SwingFXUtils.toFXImage(cardObject.getCardBufferedImages().get(indexOfROI), null);
    }

    private void saveAllCardDataToDataFile(){
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            fis.close();

            Sheet sh = wb.createSheet("Card Data");

            XSSFRow row0 = (XSSFRow) sh.createRow(0);
            row0.createCell(0).setCellValue("Card Spacing:");
            row0.createCell(1).setCellValue(CARD_SPACING);
            row0.createCell(3).setCellValue("Card Index:");

            XSSFRow row1 = (XSSFRow) sh.createRow(1);
            row1.createCell(0).setCellValue("Cards on Pass:");
            row1.createCell(1).setCellValue(Integer.parseInt(spinnerCardPassNum.getValue().toString()));
            row1.createCell(3).setCellValue("Scanned?");

            XSSFRow row2 = (XSSFRow) sh.createRow(2);
            row2.createCell(0).setCellValue("Spread Calc:");
            if (useSpreadFactor==1) {
                row2.createCell(1).setCellValue("Adaptive");
            } else {
                row2.createCell(1).setCellValue("Direct");
            }
            row2.createCell(3).setCellValue("Used?");

            XSSFRow row3 = (XSSFRow) sh.createRow(3);
            row3.createCell(0).setCellValue("Spread-A");
            row3.createCell(1).setCellValue(spreadA);
            row3.createCell(3).setCellValue("Threshold:");

            XSSFRow row4 = (XSSFRow) sh.createRow(4);
            row4.createCell(0).setCellValue("Spread-B");
            row4.createCell(1).setCellValue(spreadB);
            row4.createCell(3).setCellValue("Drop Count:");

            XSSFRow row5 = (XSSFRow) sh.createRow(5);
            row5.createCell(0).setCellValue("Spread-C");
            row5.createCell(1).setCellValue(spreadC);
            row5.createCell(3).setCellValue("% Coverage:");

            XSSFRow row6 = (XSSFRow) sh.createRow(6);
            row6.createCell(3).setCellValue("Card (\u00B5m^2):");

            XSSFRow row7 = (XSSFRow) sh.createRow(7);
            row7.createCell(3).setCellValue("Drops/in^2:");

            XSSFRow row8 = (XSSFRow) sh.createRow(8);
            row8.createCell(3).setCellValue("VMD:");

            XSSFRow row9 = (XSSFRow) sh.createRow(9);
            row9.createCell(3).setCellValue("Dv0.1:");

            XSSFRow row10 = (XSSFRow) sh.createRow(10);
            row10.createCell(3).setCellValue("Dv0.9:");

            XSSFRow row12 = (XSSFRow) sh.createRow(12);
            row12.createCell(3).setCellValue("Stains (\u00B5m^2)");

            int index = 0;
            for(int i=0; i<cardObject.getCardList().size(); i++){
                row1.createCell(4+i).setCellValue(cardObject.getCardList().get(i));
                row2.createCell(4+i).setCellValue(cardObject.getUsingCardList().get(i));
                if(cardObject.getCardList().get(i)){
                    row0.createCell(4+i).setCellValue(cardObject.getCardNameList().get(index));
                    row3.createCell(4+i).setCellValue(cardObject.getThresholds().get(index));
                    row4.createCell(4+i).setCellValue(cardObject.getStains().get(index).size());
                    row5.createCell(4+i).setCellValue(cardObject.getPercentCoverage().get(index));
                    row6.createCell(4+i).setCellValue(cardObject.getCardAreas().get(index));
                    row7.createCell(4+i).setCellValue(cardObject.getDropsPerSquareInch().get(index));
                    row8.createCell(4+i).setCellValue(cardObject.getdV05s().get(index));
                    row9.createCell(4+i).setCellValue(cardObject.getdV01s().get(index));
                    row10.createCell(4+i).setCellValue(cardObject.getdV09s().get(index));

                    for(int z=0; z<cardObject.getStains().get(index).size(); z++){
                        if(sh.getRow(z+12)!=null){
                            sh.getRow(z+12).createCell(4+i).setCellValue(cardObject.getStains().get(index).get(z));
                        } else {
                            sh.createRow(z+12).createCell(4+i).setCellValue(cardObject.getStains().get(index).get(z));
                        }
                    }

                    index++;
                }
            }

            sh.autoSizeColumn(0);
            sh.autoSizeColumn(1);
            sh.autoSizeColumn(3);

            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
        for(int i=0; i<cardObject.getCardNameList().size(); i++){
            saveCardImageToDataFile(i);
        }
    }

    private void saveCardImageToDataFile(int indexOfROI){
        //Save Pattern Data to Current Data File
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            fis.close();

            XSSFSheet sh = (XSSFSheet) wb.createSheet(cardObject.getCardNameList().get(indexOfROI));
            //cardObject.getRm().select(indexOfROI);
            //ImagePlus imp0 = cardObject.getImp().duplicate();
            //BufferedImage bufferedImage = imp0.getBufferedImage();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            BufferedImage bufferedImage = cardObject.getCardBufferedImages().get(indexOfROI);
            int width = bufferedImage.getWidth();
            int height = bufferedImage.getHeight();
            ImageIO.write(cardObject.getCardBufferedImages().get(indexOfROI), "png", baos);
            byte[] bytes = baos.toByteArray();
            int pictureIdx = wb.addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_PNG);

            XSSFCreationHelper helper = ((XSSFWorkbook) wb).getCreationHelper();

            // Create the drawing patriarch.  This is the top level container for all shapes.
            XSSFDrawing drawing = sh.createDrawingPatriarch();

            //add a picture shape
            XSSFClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.setCol1(0);
            anchor.setRow1(0);
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize();



            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();
        } catch (Exception e) {e.printStackTrace();}
    }

    private javafx.scene.image.Image thresholdedSnapshotOld(int indexOfROI, int updatedThreshold){
        ImagePlus imp0 = thresholdImage(indexOfROI, updatedThreshold);
        BufferedImage bufferedImage = imp0.getBufferedImage();
        return SwingFXUtils.toFXImage(bufferedImage, null);
    }

    private javafx.scene.image.Image thresholdedSnapshot(int indexOfROI, int updatedThreshold){
        ImagePlus imp = thresholdImage(indexOfROI, updatedThreshold);
        IJ.run(imp, "Watershed", "");
        //Analyze Particles
        ResultsTable rt = ResultsTable.getResultsTable();
        ParticleAnalyzer analyzer = new ParticleAnalyzer(ParticleAnalyzer.CLEAR_WORKSHEET+ParticleAnalyzer.INCLUDE_HOLES+ParticleAnalyzer.EXCLUDE_EDGE_PARTICLES+ParticleAnalyzer.SHOW_OVERLAY_MASKS,
                1, rt,
                1.0D,
                1.7976931348623157E308D,
                0.0D,
                1.0D);
        analyzer.analyze(imp);
        analyzer.setHideOutputImage(true);
        //IJ.run(imp, "Overlay Options", "stroke=black width=0 fill=black apply");
        //Overlay overlay = imp.getOverlay();
        if (rt.getCounter() > 0) {
            imp.getOverlay().setStrokeColor(Color.RED);
            imp.getOverlay().setFillColor(Color.BLUE);
            imp.getOverlay().drawLabels(false);
        }

        BufferedImage bufferedImage = imp.flatten().getBufferedImage();
        return SwingFXUtils.toFXImage(bufferedImage, null);
    }

    private ImagePlus thresholdImage(int indexOfROI, int updatedThreshold){
        int appropriateThreshold = updatedThreshold;
        //cardObject.getRm().select(indexOfROI);
        ImagePlus imp = new ImagePlus("From BuffIm",cardObject.getCardBufferedImages().get(indexOfROI));

        //Convert and DEFAULT_THRESHOLD
        ImageConverter ic = new ImageConverter(imp);
        ic.convertToGray8();
        imp.updateImage();
        ImageProcessor proc0 = imp.getProcessor();
        proc0.setAutoThreshold(AutoThresholder.Method.Default,false);
        if (updatedThreshold<0) {
            if (proc0.getMaxThreshold()<= DEFAULT_THRESHOLD) {
                appropriateThreshold = (int) proc0.getMaxThreshold();
                proc0.autoThreshold();
            } else {
                //proc0.setThreshold(0.0, DEFAULT_THRESHOLD, ImageProcessor.NO_LUT_UPDATE);
                proc0.threshold(DEFAULT_THRESHOLD);
                appropriateThreshold = DEFAULT_THRESHOLD;
            }
            cardObject.getThresholds().set(indexOfROI,appropriateThreshold);
        } else {
            proc0.threshold(updatedThreshold);
            //proc0.setThreshold(0.0, updatedThreshold, ImageProcessor.NO_LUT_UPDATE);
        }

        ImagePlus imp1 = new ImagePlus("thresholdedImage",proc0);
        return imp1;
    }

    private ArrayList<Double> getStainsFromCard(ImagePlus thresholdedImage){
        //Run watershed algorithm to bust up buddies
        IJ.run(thresholdedImage, "Watershed", "");
        //Analyze Particles
        ResultsTable rt = ResultsTable.getResultsTable();
        ParticleAnalyzer analyzer = new ParticleAnalyzer(ParticleAnalyzer.CLEAR_WORKSHEET+ParticleAnalyzer.INCLUDE_HOLES+ParticleAnalyzer.EXCLUDE_EDGE_PARTICLES,
                1, rt,
                1.0D,
                1.7976931348623157E308D,
                0.0D,
                1.0D);
        analyzer.setHideOutputImage(true);
        analyzer.analyze(thresholdedImage);
        rt.updateResults();

        //Stain areas First
        int dropCountCard = rt.getCounter();
        ArrayList<Double> areaStainPX = new ArrayList<>();
        for(int i=0; i<dropCountCard; i++){
            double a = rt.getValueAsDouble(0, i);
            if (a>4) {
                areaStainPX.add(a);
            }
        }
        //Convert to square micron areas
        ArrayList<Double> areaStainUM = new ArrayList<>();
        for(int i=0; i<areaStainPX.size(); i++){
            areaStainUM.add(i, areaStainPX.get(i)*42.3333D*42.3333D);
        }
        //cardObject.getStains().add(areaStainUM);
        return areaStainUM;

    }

    private void iterateThroughCards(boolean useEstablishedThresholds){
        if (useEstablishedThresholds) {
            //For opening files with pre-established thresholds
            int indexOfROI = 0;
            for(int i=0; i<cardObject.getCardList().size(); i++){
                if(cardObject.getCardList().get(i)){
                    cardObject.getCardImages().add(originalSnapshot(indexOfROI));
                    cardObject.getCardImagesThresholded().add(thresholdedSnapshot(indexOfROI, cardObject.getThresholds().get(indexOfROI)));
                    logStains(getStainsFromCard(thresholdImage(indexOfROI,cardObject.getThresholds().get(indexOfROI))),indexOfROI);
                    logPercentCoverage(indexOfROI, cardObject.getThresholds().get(indexOfROI), true);
                    displayCardOnScreen(indexOfROI,i,cardObject.getCardImages().get(indexOfROI),cardObject.getCardImagesThresholded().get(indexOfROI),cardObject.getThresholds().get(indexOfROI));

                    indexOfROI++;
                }
            }
        } else {
            //For running new data (from process full scan)
            int indexOfROI = 0;
            for(int i=0; i<cardObject.getCardList().size(); i++){
                if(cardObject.getCardList().get(i)){
                    cardObject.getCardImages().add(originalSnapshot(indexOfROI));
                    cardObject.getCardImagesThresholded().add(thresholdedSnapshot(indexOfROI, -1));
                    logStains(getStainsFromCard(thresholdImage(indexOfROI,-1)),indexOfROI);
                    logPercentCoverage(indexOfROI, -1, true);
                    displayCardOnScreen(indexOfROI,i,cardObject.getCardImages().get(indexOfROI),cardObject.getCardImagesThresholded().get(indexOfROI),cardObject.getThresholds().get(indexOfROI));

                    indexOfROI++;
                }
            }
        }
    }

    private void logStains(ArrayList<Double> stains, int indexOfROI){
        //Find Actual Diameters Using Spread Factor
        ArrayList<Double> areaStainUM = stains;
        int dropCountCard = stains.size();
        ArrayList<Double> diameterUM = new ArrayList<>();
        ArrayList<Double> volumeUM3 = new ArrayList<>();
        double areaStainUMCumulative = 0;
        double volumeCumulativeCard = 0;
        for(int i=0; i<dropCountCard; i++){
            double ds = Math.sqrt((4.0D * areaStainUM.get(i)) / Math.PI);
            double dd;
            if (useSpreadFactor==1){
                dd = ds / ((spreadA*Math.pow(ds, 2.0D)) + (spreadB*ds) + (spreadC));
            } else {
                dd = (spreadA*Math.pow(ds, 2.0D)) + (spreadB*ds) + (spreadC);
            }
            double v = (Math.PI*Math.pow(dd, 3.0D)) / 6.0D;
            diameterUM.add(dd);
            volumeUM3.add(v);
            volumeCumulativeCard+=v;
            areaStainUMCumulative+=areaStainUM.get(i);
        }
        Collections.sort(diameterUM);
        Collections.sort(volumeUM3);

        //double countVol = 0;
        ArrayList<Double> countVols = new ArrayList<>();
        boolean dv01 = false;
        boolean dv05 = false;
        boolean dv09 = false;
        if (dropCountCard>0) {
            for(int i=0; i<dropCountCard; i++){
                //Make a list of cumulative volumes
                if(i>0){
                    countVols.add(volumeUM3.get(i)+countVols.get(i-1));
                } else {
                    countVols.add(volumeUM3.get(i));
                }
                if(!dv01 && countVols.get(i) >= (volumeCumulativeCard*0.1D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.1D) {
                        cardObject.getdV01s().add(diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV01s().add(0 + ((((volumeCumulativeCard * 0.1D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV01s().add(diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.1D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));
                    }
                    dv01 = true;
                }
                if(!dv05 && countVols.get(i) >= (volumeCumulativeCard*0.5D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.5D) {
                        cardObject.getdV05s().add(diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV05s().add(0 + ((((volumeCumulativeCard * 0.5D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV05s().add(diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.5D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));                }
                    dv05 = true;
                }
                if(!dv09 && countVols.get(i) >= (volumeCumulativeCard*0.9D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.9D) {
                        cardObject.getdV09s().add(diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV09s().add(0 + ((((volumeCumulativeCard * 0.9D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV09s().add(diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.9D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));                }
                    dv09 = true;
                }
            }
        } else {
            cardObject.getdV01s().add(0.0D);
            cardObject.getdV05s().add(0.0D);
            cardObject.getdV09s().add(0.0D);
        }
        cardObject.getStains().add(stains);
        cardObject.getDiameters().add(diameterUM);
        cardObject.getVolumes().add(volumeUM3);
//        double w = cardObject.getRm().getRoi(cardObject.getRm().getSelectedIndex()).getBounds().getWidth();
//        double h = cardObject.getRm().getRoi(cardObject.getRm().getSelectedIndex()).getBounds().getHeight();
//        double cardArea = (w*h);
//        double cardAreaUM = cardArea*42.3333D*42.3333D;
        cardObject.getCardAreas().add(cardObject.getCardBufferedImages().get(indexOfROI).getWidth()*cardObject.getCardBufferedImages().get(indexOfROI).getHeight()*42.3333D*42.3333D);
        //cardObject.getPercentCoverage().add((areaStainUMCumulative/cardObject.getCardAreas().get(indexOfROI))*100);

        double cardAreaSqInch = cardObject.getCardAreas().get(indexOfROI)/645160000.0;
        cardObject.getDropsPerSquareInch().add(dropCountCard/cardAreaSqInch);
    }

    private void relogStains(ArrayList<Double> stains, int indexOfROI){
        //Find Actual Diameters Using Spread Factor
        ArrayList<Double> areaStainUM = stains;
        int dropCountCard = stains.size();
        ArrayList<Double> diameterUM = new ArrayList<>();
        ArrayList<Double> volumeUM3 = new ArrayList<>();
        double areaStainUMCumulative = 0;
        double volumeCumulativeCard = 0;
        for(int i=0; i<dropCountCard; i++){
            double ds = Math.sqrt((4.0D * areaStainUM.get(i)) / Math.PI);
            double dd;
            if (useSpreadFactor==1){
                dd = ds / ((spreadA*Math.pow(ds, 2.0D)) + (spreadB*ds) + (spreadC));
            } else {
                dd = (spreadA*Math.pow(ds, 2.0D)) + (spreadB*ds) + (spreadC);
            }
            double v = (Math.PI*Math.pow(dd, 3.0D)) / 6.0D;
            diameterUM.add(dd);
            volumeUM3.add(v);
            volumeCumulativeCard+=v;
            areaStainUMCumulative+=areaStainUM.get(i);
        }
        Collections.sort(diameterUM);
        Collections.sort(volumeUM3);

        //double countVol = 0;
        ArrayList<Double> countVols = new ArrayList<>();
        boolean dv01 = false;
        boolean dv05 = false;
        boolean dv09 = false;
        if (dropCountCard>0) {
            for(int i=0; i<dropCountCard; i++){
                //Make a list of cumulative volumes
                if(i>0){
                    countVols.add(volumeUM3.get(i)+countVols.get(i-1));
                } else {
                    countVols.add(volumeUM3.get(i));
                }
                if(!dv01 && countVols.get(i) >= (volumeCumulativeCard*0.1D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.1D) {
                        cardObject.getdV01s().set(indexOfROI,diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV01s().set(indexOfROI,0 + ((((volumeCumulativeCard * 0.1D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV01s().set(indexOfROI,diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.1D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));
                    }
                    dv01 = true;
                }
                if(!dv05 && countVols.get(i) >= (volumeCumulativeCard*0.5D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.5D) {
                        cardObject.getdV05s().set(indexOfROI,diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV05s().set(indexOfROI,0 + ((((volumeCumulativeCard * 0.5D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV05s().set(indexOfROI,diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.5D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));                }
                    dv05 = true;
                }
                if(!dv09 && countVols.get(i) >= (volumeCumulativeCard*0.9D)){
                    if (countVols.get(i) == volumeCumulativeCard*0.9D) {
                        cardObject.getdV09s().set(indexOfROI,diameterUM.get(i));
                    } else if (i==0) {
                        cardObject.getdV09s().set(indexOfROI,0 + ((((volumeCumulativeCard * 0.9D) - 0) / (countVols.get(i) - 0)) * (diameterUM.get(i) - 0)));
                    } else {
                        cardObject.getdV09s().set(indexOfROI,diameterUM.get(i - 1) + ((((volumeCumulativeCard * 0.9D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterUM.get(i) - diameterUM.get(i - 1))));                }
                    dv09 = true;
                }
            }
        } else {
            cardObject.getdV01s().set(indexOfROI,0.0D);
            cardObject.getdV05s().set(indexOfROI,0.0D);
            cardObject.getdV09s().set(indexOfROI,0.0D);
        }
        cardObject.getStains().set(indexOfROI,stains);
        cardObject.getDiameters().set(indexOfROI,diameterUM);
        cardObject.getVolumes().set(indexOfROI,volumeUM3);
        //cardObject.getCardAreas().set(indexOfROI,cardObject.getCardBufferedImages().get(indexOfROI).getWidth()*cardObject.getCardBufferedImages().get(indexOfROI).getHeight()*42.3333D*42.3333D);
        //cardObject.getPercentCoverage().set(indexOfROI,(areaStainUMCumulative/cardObject.getCardAreas().get(indexOfROI))*100);

        double cardAreaSqInch = cardObject.getCardAreas().get(indexOfROI)/645160000.0;
        cardObject.getDropsPerSquareInch().set(indexOfROI,dropCountCard/cardAreaSqInch);
    }

    private void logPercentCoverage(int indexOfROI, int updatedThreshold, boolean isNew){
        ImagePlus imp = thresholdImage(indexOfROI, updatedThreshold);
        IJ.run(imp, "Watershed", "");

        //Set Conversion factor based on dpi
        double micronsPerPixel = 25400 / FULLSCAN_DPI;

        //Analyze Particles Including Edges (For % Coverage)
        ResultsTable rt0 = ResultsTable.getResultsTable();
        ParticleAnalyzer analyzer0 = new ParticleAnalyzer(ParticleAnalyzer.CLEAR_WORKSHEET+ParticleAnalyzer.INCLUDE_HOLES,
                1, rt0,
                1.0D,
                1.7976931348623157E308D,
                0.0D,
                1.0D);
        analyzer0.setHideOutputImage(true);
        analyzer0.analyze(imp);
        rt0.updateResults();

        //Get stain areas, exclude areas less than specified
        double sumCoverage = 0.0;
        for(int i=0; i<rt0.getCounter(); i++){
            double area = rt0.getValueAsDouble(0,i);
            if(area >= 4){
                sumCoverage += (area * micronsPerPixel * micronsPerPixel);
            }
        }
        if (isNew) {
            cardObject.getPercentCoverage().add((sumCoverage/cardObject.getCardAreas().get(indexOfROI))*100);
        } else {
            cardObject.getPercentCoverage().set(indexOfROI, (sumCoverage/cardObject.getCardAreas().get(indexOfROI))*100);
        }
    }

    private void showComposites(){
        computeComposites();
        plotCoverage();
        plotDropletDist();
    }

    private void computeComposites(){
        ArrayList<Double> volumeList = new ArrayList<>();
        ArrayList<Double> diameterList = new ArrayList<>();
        int index = 0;
        for(int i=0; i<cardObject.getUsingCardList().size(); i++){
            if (cardObject.getCardList().get(i)) {
                if (cardObject.getUsingCardList().get(i)) {
                    volumeList.addAll(cardObject.getVolumes().get(index));
                    diameterList.addAll(cardObject.getDiameters().get(index));
                }
                index++;
            }
        }
        Collections.sort(volumeList);
        Collections.sort(diameterList);
        double volumeSum = 0;
        double dropCount = 0;
        for(int i=0; i<volumeList.size(); i++){
            volumeSum+=volumeList.get(i);
            dropCount++;
        }


        //Composite Data
        boolean dv01 = false;
        double dV01 = 0;
        boolean dv05 = false;
        double dV05 = 0;
        boolean dv09 = false;
        double dV09 = 0;
        ArrayList<Double> countVols = new ArrayList<>();
        for (int i = 0; i < dropCount; i++) {
            //countVol += volumeUM3.get(i);
            //Make a list of cumulative volumes
            if (i > 0) {
                countVols.add(volumeList.get(i) + countVols.get(i - 1));
            } else {
                countVols.add(volumeList.get(i));
            }
            if (!dv01 && countVols.get(i) >= (volumeSum * 0.1D)) {
                if (countVols.get(i) == volumeSum * 0.1D) {
                    dV01 = diameterList.get(i);
                } else if (i == 0) {
                    dV01 = 0 + ((((volumeSum * 0.1D) - 0) / (countVols.get(i) - 0)) * (diameterList.get(i) - 0));

                } else {
                    dV01 = diameterList.get(i - 1) + ((((volumeSum * 0.1D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterList.get(i) - diameterList.get(i - 1)));
                }
                dv01 = true;
            }
            if (!dv05 && countVols.get(i) >= (volumeSum * 0.5D)) {
                if (countVols.get(i) == volumeSum * 0.5D) {
                    dV05 = diameterList.get(i);
                } else if (i == 0) {
                    dV05 = 0 + ((((volumeSum * 0.5D) - 0) / (countVols.get(i) - 0)) * (diameterList.get(i) - 0));

                } else {
                    dV05 = diameterList.get(i - 1) + ((((volumeSum * 0.5D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterList.get(i) - diameterList.get(i - 1)));
                }
                dv05 = true;
            }
            if (!dv09 && countVols.get(i) >= (volumeSum * 0.9D)) {
                if (countVols.get(i) == volumeSum * 0.9D) {
                    dV09 = diameterList.get(i);
                } else if (i == 0) {
                    dV09 = 0 + ((((volumeSum * 0.9D) - 0) / (countVols.get(i) - 0)) * (diameterList.get(i) - 0));

                } else {
                    dV09 = diameterList.get(i - 1) + ((((volumeSum * 0.9D) - countVols.get(i - 1)) / (countVols.get(i) - countVols.get(i - 1))) * (diameterList.get(i) - diameterList.get(i - 1)));
                }
                dv09 = true;
            }
        }

        double RS = (dV09 - dV01) / dV05;

        double totCov = 0;
        double totCovPts = 0;
        int indexx = 0;
        for(int i=0; i<cardObject.getUsingCardList().size(); i++){
            if (cardObject.getCardList().get(i)) {
                if (cardObject.getUsingCardList().get(i)) {
                    totCov += cardObject.getPercentCoverage().get(indexx);
                    totCovPts++;
                }
                indexx++;
            }
        }
        labelCCov.setText(String.format("%.2f", totCov / totCovPts) + "%");
        labelCVMD.setText(Integer.toString((int) dV05) + " \u00B5m");
        labelCDV01.setText(Integer.toString((int) dV01) + " \u00B5m");
        labelCDV09.setText(Integer.toString((int) dV09) + " \u00B5m");
        labelCRS.setText(String.format("%.2f",RS));
        labelCDSC.setText(ModelCalc.getDSC(dV01,dV05));
    }

    private void processFullScan(File fullScanFile){
        Opener opener = new Opener();
        cardObject.setImp(opener.openImage(fullScanFile.getPath()));

        //Testing of adding rois
        if(cardObject.getOverlay().size()!=cardObject.getRm().getCount()){
            cardObject.getOverlay().clear();
            for(int i=0; i<cardObject.getRm().getCount(); i++){
                cardObject.getOverlay().add(cardObject.getRm().getRoi(i));
            }
        }
        //
        cardObject.getImp().show();
        cardObject.getImp().setOverlay(cardObject.getOverlay());
//        ip = new ImagePlus();
//        ip = imp.duplicate();
//        ip.show();
//        ip.setOverlay(overlay);
//        //generateROIS();


        //Delete smaller ROIS, rescale overlay boxes and add as new ROIs to ROIManager
        cardObject.getRm().runCommand("Select All");
        cardObject.getRm().runCommand("Delete");
        for(int i=0; i<cardObject.getOverlay().size(); i++){
            java.awt.Rectangle rectOG = cardObject.getOverlay().get(i).getBounds();
            int q = FULLSCAN_DPI/PREVIEW_DPI;
            int xcoord = rectOG.x*q;
            int ycoord = rectOG.y*q;
            int width = rectOG.width*q;
            int height = rectOG.height*q;
            java.awt.Rectangle rectNEW = new java.awt.Rectangle(xcoord,ycoord,width,height);
            cardObject.getRm().addRoi(new Roi(rectNEW));
            cardObject.getRm().select(i);
            cardObject.getCardBufferedImages().add(cardObject.getImp().duplicate().getBufferedImage());
            //System.out.println("From rectOG Image "+i+" - Height: "+height+" Width: "+width);
            //cardObject.getCardAreas().add(width*height*42.3333D*42.3333D);
        }
        cardObject.getImp().getWindow().setVisible(false);

        //Send Rois to CardObject as BufferedImages
//        for(int i=0; i<cardObject.getRm().getCount(); i++){
//            cardObject.getRm().select(i);cardObject.getCardAreas().add(cardObject.getCardBufferedImages().get(indexOfROI).getWidth()*cardObject.getCardBufferedImages().get(indexOfROI).getHeight()*42.3333D*42.3333D);
//            cardObject.getCardBufferedImages().add(cardObject.getImp().duplicate().getBufferedImage());
//            int w = cardObject.getCardBufferedImages().get(i).getWidth();
//            int h = cardObject.getCardBufferedImages().get(i).getHeight();
//            System.out.println("From Buffered Image "+i+" - Height: "+h+" Width: "+w);
//        }

        assignScannedCards();
        iterateThroughCards(false);
        saveAllCardDataToDataFile();
        showComposites();
    }

    private void displayCardOnScreen(int index, int card, javafx.scene.image.Image img, javafx.scene.image.Image imgT, int cardThreshold){

        switch (card) {
            case 0:
                labelCard0.setText(cardObject.getCardNameList().get(index));
                imageViewCard0.setImage(img);
                labelCoverage0.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD0.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops0.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard0T.setImage(imgT);
                labelCard0T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 1:
                labelCard1.setText(cardObject.getCardNameList().get(index));
                imageViewCard1.setImage(img);
                labelCoverage1.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD1.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops1.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard1T.setImage(imgT);
                labelCard1T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 2:
                labelCard2.setText(cardObject.getCardNameList().get(index));
                imageViewCard2.setImage(img);
                labelCoverage2.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD2.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops2.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard2T.setImage(imgT);
                labelCard2T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 3:
                labelCard3.setText(cardObject.getCardNameList().get(index));
                imageViewCard3.setImage(img);
                labelCoverage3.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD3.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops3.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard3T.setImage(imgT);
                labelCard3T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 4:
                labelCard4.setText(cardObject.getCardNameList().get(index));
                imageViewCard4.setImage(img);
                labelCoverage4.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD4.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops4.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard4T.setImage(imgT);
                labelCard4T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 5:
                labelCard5.setText(cardObject.getCardNameList().get(index));
                imageViewCard5.setImage(img);
                labelCoverage5.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD5.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops5.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard5T.setImage(imgT);
                labelCard5T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 6:
                labelCard6.setText(cardObject.getCardNameList().get(index));
                imageViewCard6.setImage(img);
                labelCoverage6.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD6.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops6.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard6T.setImage(imgT);
                labelCard6T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 7:
                labelCard7.setText(cardObject.getCardNameList().get(index));
                imageViewCard7.setImage(img);
                labelCoverage7.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD7.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops7.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard7T.setImage(imgT);
                labelCard7T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
            case 8:
                labelCard8.setText(cardObject.getCardNameList().get(index));
                imageViewCard8.setImage(img);
                labelCoverage8.setText(String.format("%.2f",cardObject.getPercentCoverage().get(index))+" % Coverage");
                labelVMD8.setText("VMD: " + Integer.toString(cardObject.getdV05s().get(index).intValue()));
                labelDrops8.setText("Drops on Card: " + Integer.toString(cardObject.getStains().get(index).size()));
                imageViewCard8T.setImage(imgT);
                labelCard8T.setText("Threshold:"+Integer.toString(cardThreshold));
                break;
        }
    }

    private void displayParticleMask(int absoluteCardIndex){
        if (cardObject.getCardList().get(absoluteCardIndex)) {
            int index = getPositionInRM(absoluteCardIndex);
            //OG image
            ImagePlus imp0 = new ImagePlus("Particle Mask for "+cardObject.getCardNameList().get(index),cardObject.getCardBufferedImages().get(index));
            //Binary thresholded
            ImagePlus imp = thresholdImage(index, cardObject.getThresholds().get(index));
            //New Imp which will be inverted
            IJ.run(imp, "Watershed", "");

            //Analyze Particles
            ResultsTable rt = ResultsTable.getResultsTable();
            ParticleAnalyzer analyzer = new ParticleAnalyzer(ParticleAnalyzer.CLEAR_WORKSHEET+ParticleAnalyzer.INCLUDE_HOLES+ParticleAnalyzer.EXCLUDE_EDGE_PARTICLES+ParticleAnalyzer.SHOW_OVERLAY_OUTLINES,
                    1, rt,
                    1.0D,
                    1.7976931348623157E308D,
                    0.0D,
                    1.0D);
            analyzer.analyze(imp);
            analyzer.setHideOutputImage(true);

            imp0.show();
            Overlay overlay = imp.getOverlay();
            overlay.drawLabels(false);
            imp0.setOverlay(overlay);
        }
    }

    private int getPositionInRM(int absoluteIndex){
        int rmIndex = 0;
        int position = 0;
        for(int i=0; i<cardObject.getCardList().size(); i++){
            if(i==absoluteIndex){
                position = rmIndex;
            }
            if(cardObject.getCardList().get(i)){
                rmIndex++;
            }
        }
        //System.out.println("getPositionInRM: absolute index = "+absoluteIndex+" position = "+position+" cardListSize = "+cardObject.getCardList().size());
        return position;
    }

    private int getAbsoluteIndex(int indexOfROI){
        int rmIndex = 0;
        int position = 0;
        for(int i=0; i<cardObject.getCardList().size(); i++){
            if(cardObject.getCardList().get(i)){
                if(rmIndex==indexOfROI){
                    position = i;
                }
                rmIndex++;
            }
        }
        System.out.println("getAbsoluteIndex: indexOfROI = "+indexOfROI+" position = "+position);
        return position;
    }

    private void rerunUpdatedCard(int absoluteIndex, int deltaThreshold){
//        if(cardObject.getImp()==null){
//            return; //Catch for opening past files
//        }
        int indexOfROI = getPositionInRM(absoluteIndex);
        cardObject.getThresholds().set(indexOfROI,cardObject.getThresholds().get(indexOfROI)+deltaThreshold);
        cardObject.getCardImagesThresholded().set(indexOfROI, thresholdedSnapshot(indexOfROI,cardObject.getThresholds().get(indexOfROI)));
        relogStains(getStainsFromCard(thresholdImage(indexOfROI, cardObject.getThresholds().get(indexOfROI))),indexOfROI);
        logPercentCoverage(indexOfROI, cardObject.getThresholds().get(indexOfROI), false);
        displayCardOnScreen(indexOfROI,absoluteIndex,cardObject.getCardImages().get(indexOfROI),cardObject.getCardImagesThresholded().get(indexOfROI),cardObject.getThresholds().get(indexOfROI));
        showComposites();

        replaceCardDataInDataFile(absoluteIndex);
    }

    private void replaceCardDataInDataFile(int absoluteIndex){
        try {
            FileInputStream fis = new FileInputStream(currentDataFile);
            Workbook wb = WorkbookFactory.create(fis);
            fis.close();

            Sheet sh = wb.getSheet("Card Data");

            int i = absoluteIndex;
            int index = getPositionInRM(absoluteIndex);
            sh.getRow(2).getCell(4+i).setCellValue(cardObject.getUsingCardList().get(i));
            if(cardObject.getCardList().get(i)){
                sh.getRow(3).getCell(4+i).setCellValue(cardObject.getThresholds().get(index));
                sh.getRow(4).getCell(4+i).setCellValue(cardObject.getStains().get(index).size());
                sh.getRow(5).getCell(4+i).setCellValue(cardObject.getPercentCoverage().get(index));
                sh.getRow(6).getCell(4+i).setCellValue(cardObject.getCardAreas().get(index));
                sh.getRow(7).getCell(4+i).setCellValue(cardObject.getDropsPerSquareInch().get(index));
                sh.getRow(8).getCell(4+i).setCellValue(cardObject.getdV05s().get(index));
                sh.getRow(9).getCell(4+i).setCellValue(cardObject.getdV01s().get(index));
                sh.getRow(10).getCell(4+i).setCellValue(cardObject.getdV09s().get(index));

                for (java.util.Iterator<Row> z = sh.rowIterator(); z.hasNext();) {
                    int row = z.next().getRowNum();
                    if (row > 11 && sh.getRow(row)!=null && sh.getRow(row).getCell(4+i)!=null){
                        sh.getRow(row).getCell(4+i).setCellType(Cell.CELL_TYPE_BLANK);
                    }
                }

                for(int z=0; z<cardObject.getStains().get(index).size(); z++){
                    if(sh.getRow(z+12)!=null){
                        if (sh.getRow(z+12).getCell(4+i)!=null) {
                            sh.getRow(z+12).getCell(4+i).setCellValue(cardObject.getStains().get(index).get(z));
                        } else {
                            sh.getRow(z+12).createCell(4+i).setCellValue(cardObject.getStains().get(index).get(z));
                        }
                    } else {
                        sh.createRow(z+12).createCell(4+i).setCellValue(cardObject.getStains().get(index).get(z));
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream(currentDataFile);
            wb.write(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void clickCard0(){
        displayParticleMask(0);
    }

    public void clickCard1(){
        displayParticleMask(1);
    }

    public void clickCard2(){
        displayParticleMask(2);
    }

    public void clickCard3(){
        displayParticleMask(3);
    }

    public void clickCard4(){
        displayParticleMask(4);
    }

    public void clickCard5(){
        displayParticleMask(5);
    }

    public void clickCard6(){
        displayParticleMask(6);
    }

    public void clickCard7(){
        displayParticleMask(7);
    }

    public void clickCard8(){
        displayParticleMask(8);
    }

    public void clickButton0TMinus(){
        rerunUpdatedCard(0,-1);
    }

    public void clickButton0TPlus(){
        rerunUpdatedCard(0,+1);
    }

    public void clickButton1TMinus(){
        rerunUpdatedCard(1,-1);
    }

    public void clickButton1TPlus(){
        rerunUpdatedCard(1,+1);
    }

    public void clickButton2TMinus(){
        rerunUpdatedCard(2,-1);
    }

    public void clickButton2TPlus(){
        rerunUpdatedCard(2,+1);
    }

    public void clickButton3TMinus(){
        rerunUpdatedCard(3,-1);
    }

    public void clickButton3TPlus(){
        rerunUpdatedCard(3,+1);
    }

    public void clickButton4TMinus(){
        rerunUpdatedCard(4,-1);
    }

    public void clickButton4TPlus(){
        rerunUpdatedCard(4,+1);
    }

    public void clickButton5TMinus(){
        rerunUpdatedCard(5,-1);
    }

    public void clickButton5TPlus(){
        rerunUpdatedCard(5,+1);
    }

    public void clickButton6TMinus(){
        rerunUpdatedCard(6,-1);
    }

    public void clickButton6TPlus(){
        rerunUpdatedCard(6,+1);
    }

    public void clickButton7TMinus(){
        rerunUpdatedCard(7,-1);
    }

    public void clickButton7TPlus(){
        rerunUpdatedCard(7,+1);
    }

    public void clickButton8TMinus(){
        rerunUpdatedCard(8,-1);
    }

    public void clickButton8TPlus(){
        rerunUpdatedCard(8,+1);
    }

    private void plotCoverage(){
        int s = CARD_SPACING;

        areaChartCoverage.setLegendVisible(false);
        areaChartCoverage.getData().clear();
        XYChart.Series series = new XYChart.Series();
        coverageXAxis.tickLabelFontProperty().set(javafx.scene.text.Font.font(14));
        NumberAxis xAxis = (NumberAxis) areaChartCoverage.getXAxis();
        xAxis.setLowerBound(-s*4);
        xAxis.setUpperBound(s*4);
        coverageYAxis.tickLabelFontProperty().set(javafx.scene.text.Font.font(14));
        int index = 0;
        int xpos = s*-4;
        for(int i =0; i<cardObject.getCardList().size(); i++){
            if (cardObject.getCardList().get(i)) {
                if (cardObject.getUsingCardList().get(i)) {
                    series.getData().add(new AreaChart.Data<Number,Number>(xpos, cardObject.getPercentCoverage().get(index)));
                }
                index++;
            }
            xpos += s;
        }
        areaChartCoverage.getData().addAll(series);
    }

    private void plotDropletDist(){
        barChartDropletVolume.setLegendVisible(false);
        barChartDropletVolume.getData().clear();
        histXAxis.tickLabelFontProperty().set(javafx.scene.text.Font.font(14));
        histYAxis.tickLabelFontProperty().set(javafx.scene.text.Font.font(14));
        XYChart.Series series1 = new XYChart.Series();

        double bin1 = 0;
        double bin2 = 0;
        double bin3 = 0;
        double bin4 = 0;
        double bin5 = 0;
        double bin6 = 0;
        double bin7 = 0;
        double bin8 = 0;
        double bin9 = 0;
        double bin10 = 0;
        double bin11 = 0;
        double bin12 = 0;
        double bin13 = 0;
        double bin14 = 0;
        double bin15 = 0;
        double bin16 = 0;
        double bin17 = 0;
        double bin18 = 0;

        ArrayList<Double> volumeList = new ArrayList<>();
        ArrayList<Double> diameterList = new ArrayList<>();
        int index = 0;
        for(int i=0; i<cardObject.getUsingCardList().size(); i++){
            if (cardObject.getCardList().get(i)) {
                if (cardObject.getUsingCardList().get(i)) {
                    volumeList.addAll(cardObject.getVolumes().get(index));
                    diameterList.addAll(cardObject.getDiameters().get(index));
                }
                index++;
            }
        }
        Collections.sort(volumeList);
        Collections.sort(diameterList);
        double volumeSum = 0;
        double dropCount = 0;
        for(int i=0; i<volumeList.size(); i++){
            volumeSum+=volumeList.get(i);
            dropCount++;
        }

        for(int i=0; i<dropCount; i++){
            //double a = areaList.get(i);
            //double dia = 2.0D * Math.sqrt(a / Math.PI);
            //double vol = (4.0D/3.0D)*Math.PI*Math.pow((dia/2.0D),3.0D);
            double dia = diameterList.get(i);
            double vol = volumeList.get(i);
            //System.out.println("area="+a+" dia="+dia);
            if (dia < 50) {
                bin1+=(vol/volumeSum)*100;
            } else if (dia >= 50 && dia < 100) {
                bin2+=(vol/volumeSum)*100;
            } else if (dia >= 100 && dia < 150) {
                bin3+=(vol/volumeSum)*100;
            } else if (dia >= 150 && dia < 200) {
                bin4+=(vol/volumeSum)*100;
            } else if (dia >= 200 && dia < 250) {
                bin5+=(vol/volumeSum)*100;
            } else if (dia >= 250 && dia < 300) {
                bin6+=(vol/volumeSum)*100;
            } else if (dia >= 300 && dia < 350) {
                bin7+=(vol/volumeSum)*100;
            } else if (dia >= 350 && dia < 400) {
                bin8+=(vol/volumeSum)*100;
            } else if (dia >= 400 && dia < 450) {
                bin9+=(vol/volumeSum)*100;
            } else if (dia >= 450 && dia < 500) {
                bin10+=(vol/volumeSum)*100;
            } else if (dia >= 500 && dia < 550) {
                bin11+=(vol/volumeSum)*100;
            } else if (dia >= 550 && dia < 600) {
                bin12+=(vol/volumeSum)*100;
            } else if (dia >= 600 && dia < 650) {
                bin13+=(vol/volumeSum)*100;
            } else if (dia >= 650 && dia < 700) {
                bin14+=(vol/volumeSum)*100;
            } else if (dia >= 700 && dia < 750) {
                bin15+=(vol/volumeSum)*100;
            } else if (dia >= 750 && dia < 800){
                bin16+=(vol/volumeSum)*100;
            } else if (dia >= 800 && dia < 850){
                bin17+=(vol/volumeSum)*100;
            } else if (dia >= 850 && dia < 900){
                bin18+=(vol/volumeSum)*100;
            }
        }
        //series1.getData().add(new XYChart.Data("<50", bin1));
        series1.getData().add(new XYChart.Data("50/100", bin2));
        series1.getData().add(new XYChart.Data("100/150", bin3));
        series1.getData().add(new XYChart.Data("150/200", bin4));
        series1.getData().add(new XYChart.Data("200/250", bin5));
        series1.getData().add(new XYChart.Data("250/300", bin6));
        series1.getData().add(new XYChart.Data("300/350", bin7));
        series1.getData().add(new XYChart.Data("350/400", bin8));
        series1.getData().add(new XYChart.Data("400/450", bin9));
        series1.getData().add(new XYChart.Data("450/500", bin10));
        series1.getData().add(new XYChart.Data("500/550", bin11));
        series1.getData().add(new XYChart.Data("550/600", bin12));
        series1.getData().add(new XYChart.Data("600/650", bin13));
        series1.getData().add(new XYChart.Data("650/700", bin14));
        series1.getData().add(new XYChart.Data("700/750", bin15));
        series1.getData().add(new XYChart.Data("750/800", bin16));
        series1.getData().add(new XYChart.Data("800/850", bin17));
        series1.getData().add(new XYChart.Data("850/900", bin18));

        barChartDropletVolume.getData().addAll(series1);
    }

    @Override
    public void transferProgress(int i) {

        Platform.runLater(new Runnable() {
            @Override
            public void run(){
                float p = (float) i/100;
                if (i==0) {
                    progressBarScan.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
                } else {
                    progressBarScan.setProgress(p);
                }
                if(i<100){
                    buttonPreview.setDisable(true);
                    buttonFullScan.setDisable(true);
                }
            }
        });
    }

    @Override
    public void transferDone(File file) {
        if (!fullScan) {
            buttonPreview.setDisable(false);
            buttonFullScan.setDisable(false);
            showPreviewScan(file);
//            Opener opener = new Opener();
//            imp = opener.openImage(file.getPath());
//            ip = generateROIS(imp);
//            ip.show();
//            ip.getWindow().toFront();

        } else {
            buttonPreview.setDisable(false);
            buttonFullScan.setDisable(false);
            //Opener opener = new Opener();
            //imp = opener.openImage(file.getPath());
            Platform.runLater(new Runnable() {
                @Override
                public void run(){
                    processFullScan(file);
                }
            });
        }
    }

    @Override
    public void transferFailed(int i, String s) {
        System.out.println("Failed");
        buttonPreview.setDisable(false);
        buttonFullScan.setDisable(false);
    }

    //Calculations
    private double[] savitskyGolayCoefficients(int nl, int nr, int order){
        if( nl + nr + 1<= order)
            throw new java.lang.IllegalArgumentException(" The order of polynomial cannot exceed the number of points being used." +
                    "If they are equal the input equals the output.");
        int N = nr + nl + 1;
        double[] ret_values = new double[N];
        double[] xvalues = new double[N];
        for(int i = 0; i<N; i++){

            xvalues[i] = -nl + i;

        }

        int counts = 2*order+1;
        double[] moments = new double[counts];
        for(int i = 0; i<counts; i++){
            for(int j = 0; j<N; j++){

                moments[i] += Math.pow(xvalues[j],i);

            }

            moments[i] = moments[i]/N;
        }



        double[][] matrix = new double[order+1][order+1];

        for(int i = 0; i<order+1; i++){
            for(int j = 0; j<order+1; j++){
                matrix[i][j] = moments[counts - i - j - 1];
            }
            System.out.println("");
        }

        Matrix A = new Matrix(matrix);

        LUDecomposition lu = A.lu();

        Matrix x = new Matrix(new double[order+1],order+1);
        Matrix y;
        double[] polynomial;

        for(int i = 0; i<N; i++){

            for(int j = 0; j<order+1; j++)
                x.set( j , 0, Math.pow(xvalues[i],order - j));

            y = lu.solve(x);

            polynomial = y.getColumnPackedCopy();
            ret_values[i] = evaluatePolynomial(polynomial, xvalues[nl])/N;
        }


        return ret_values;
    }

    private double savitskyGolayFilter(double[] o, double[] mask, int start, int middle){
        double out = 0;
        for(int i = 0; i<mask.length; i++)
            out += mask[i]*o[start - middle + i];

        return out;

    }

    private double reflectiveSavitskyGolayFilter(double[] o, double[] mask, int start, int middle){
        double out = 0;
        int dex;
        for(int i = 0; i<mask.length; i++){
            dex = Math.abs(start - middle + 1 +i);
            dex = (dex<o.length)?dex: 2*o.length - dex - 1;
            out += mask[i]*o[dex];
        }

        return out;

    }

    private double evaluatePolynomial(double[] poly, double x){

        double val = 0;
        int m = poly.length;
        for(int j = 0; j<m; j++){
            val += Math.pow(x,m-j-1)*poly[j];
        }
        return val;
    }

    public int airspeedCalc(int pass){
        double airspeed = 0;
        double gS = 0;
        double pH = 0;
        double wS = 0;
        double wD = 0;

        if(pass==1 && isScannedPass1){
            TextField groundSpeed = textFieldGS1;
            TextField passHeading = textFieldPH1;
            TextField windSpeed = textFieldWV1;
            TextField windDirection = textFieldWD1;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        } else if(pass==2 && isScannedPass2){
            TextField groundSpeed = textFieldGS2;
            TextField passHeading = textFieldPH2;
            TextField windSpeed = textFieldWV2;
            TextField windDirection = textFieldWD2;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        } else if(pass==3 && isScannedPass3){
            TextField groundSpeed = textFieldGS3;
            TextField passHeading = textFieldPH3;
            TextField windSpeed = textFieldWV3;
            TextField windDirection = textFieldWD3;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        } else if(pass==4 && isScannedPass4){
            TextField groundSpeed = textFieldGS4;
            TextField passHeading = textFieldPH4;
            TextField windSpeed = textFieldWV4;
            TextField windDirection = textFieldWD4;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        } else if(pass==5 && isScannedPass5){
            TextField groundSpeed = textFieldGS5;
            TextField passHeading = textFieldPH5;
            TextField windSpeed = textFieldWV5;
            TextField windDirection = textFieldWD5;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        } else if(pass==6 && isScannedPass6){
            TextField groundSpeed = textFieldGS6;
            TextField passHeading = textFieldPH6;
            TextField windSpeed = textFieldWV6;
            TextField windDirection = textFieldWD6;
            if (!groundSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windSpeed.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                gS = Double.parseDouble(groundSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wS = Double.parseDouble(windSpeed.getText());
                wD = Double.parseDouble(windDirection.getText());
                /*if(METRIC){
                    wS = wS*0.621371;
                }
                airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));*/
            }
        }
        if(invertPH){
            pH = pH - 180;
        }
        if(METRIC){
            wS = wS*0.621371;
        }
        airspeed = gS - (wS * Math.cos(Math.toRadians(wD - pH)));
        return (int) round(airspeed, 0);
    }

    public double getSprayHeight(int pass){
        if(pass==1 && isScannedPass1){
            return Double.parseDouble(textFieldSH1.getText());
        } else if(pass==2 && isScannedPass2){
            return Double.parseDouble(textFieldSH2.getText());
        } else if(pass==3 && isScannedPass3){
            return Double.parseDouble(textFieldSH3.getText());
        } else if(pass==4 && isScannedPass4){
            return Double.parseDouble(textFieldSH4.getText());
        } else if(pass==5 && isScannedPass5){
            return Double.parseDouble(textFieldSH5.getText());
        } else if(pass==6 && isScannedPass6){
            return Double.parseDouble(textFieldSH6.getText());
        } else {
            return 0;
        }
    }

    public double getWindVel(int pass){
        if(pass==1 && isScannedPass1){
            return Double.parseDouble(textFieldWV1.getText());
        } else if(pass==2 && isScannedPass2){
            return Double.parseDouble(textFieldWV2.getText());
        } else if(pass==3 && isScannedPass3){
            return Double.parseDouble(textFieldWV3.getText());
        } else if(pass==4 && isScannedPass4){
            return Double.parseDouble(textFieldWV4.getText());
        } else if(pass==5 && isScannedPass5){
            return Double.parseDouble(textFieldWV5.getText());
        } else if(pass==6 && isScannedPass6){
            return Double.parseDouble(textFieldWV6.getText());
        } else {
            return 0;
        }
    }

    public double crossWindCalc(int pass){
        double crossWind = 0;
        double wS = 0;
        double pH = 0;
        double wD = 0;

        if(pass==1 && isScannedPass1){
            TextField windSpeed = textFieldWV1;
            TextField passHeading = textFieldPH1;
            TextField windDirection = textFieldWD1;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        } else if (pass==2 && isScannedPass2){
            TextField windSpeed = textFieldWV2;
            TextField passHeading = textFieldPH2;
            TextField windDirection = textFieldWD2;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        } else if (pass==3 && isScannedPass3){
            TextField windSpeed = textFieldWV3;
            TextField passHeading = textFieldPH3;
            TextField windDirection = textFieldWD3;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        } else if (pass==4 && isScannedPass4){
            TextField windSpeed = textFieldWV4;
            TextField passHeading = textFieldPH4;
            TextField windDirection = textFieldWD4;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        } else if (pass==5 && isScannedPass5){
            TextField windSpeed = textFieldWV5;
            TextField passHeading = textFieldPH5;
            TextField windDirection = textFieldWD5;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        } else if (pass==6 && isScannedPass6){
            TextField windSpeed = textFieldWV6;
            TextField passHeading = textFieldPH6;
            TextField windDirection = textFieldWD6;
            if(!windSpeed.getText().isEmpty() && !passHeading.getText().isEmpty() && !windDirection.getText().isEmpty()) {
                wS = Double.parseDouble(windSpeed.getText());
                pH = Double.parseDouble(passHeading.getText());
                wD = Double.parseDouble(windDirection.getText());
            }
        }
        if(invertPH){
            pH = pH - 180;
        }
        crossWind = wS * Math.sin(Math.toRadians(pH - wD));
        return round(crossWind,1);
    }

    public static double round (double value, int precision) {
        int scale = (int) Math.pow(10, precision);
        return (double) Math.round(value * scale) / scale;
    }

    public double averageD (double valueA, double valueB, double valueC, double valueD, double  valueE, double valueF, int precision){
        double average = 0;
        double div = 0;
        if(isScannedPass1){
            average += valueA;
            div += 1;
        }
        if(isScannedPass2){
            average += valueB;
            div += 1;
        }
        if(isScannedPass3){
            average += valueC;
            div += 1;
        }
        if(isScannedPass4){
            average += valueD;
            div += 1;
        }
        if(isScannedPass5){
            average += valueE;
            div += 1;
        }
        if(isScannedPass6){
            average += valueF;
            div += 1;
        }
        double ave = average / div;
        return round(ave, precision);
    }

    public int averageI (int valueA, int valueB, int valueC, int valueD, int valueE, int valueF, int precision){
        double average = 0;
        double div = 0;
        if(isScannedPass1){
            average += valueA;
            div += 1;
        }
        if(isScannedPass2){
            average += valueB;
            div += 1;
        }
        if(isScannedPass3){
            average += valueC;
            div += 1;
        }
        if(isScannedPass4){
            average += valueD;
            div += 1;
        }
        if(isScannedPass5){
            average += valueE;
            div += 1;
        }
        if(isScannedPass6){
            average += valueF;
            div += 1;
        }
        double ave = average / div;
        return (int) round(ave, precision);
    }

    public double calcEstGPM (){
        double PSI= Double.parseDouble(textFieldBoomPressure.getText());
        double GPM1 = 0;
        double GPM2 = 0;
        double GPM = 0;
        double orif = 0;
        int orifPos = 0;
        if (!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            try{
                FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                if(sh==null){
                    return -9;
                }
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
                GPM1 = sh.getRow(51).getCell(orifPos).getNumericCellValue();

                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
            try{
                FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet(comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString());
                if(sh==null){
                    return -9;
                }
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
                GPM2 = sh.getRow(51).getCell(orifPos).getNumericCellValue();

                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
            int numNozzle1 = Integer.parseInt(textFieldNozzle1Quant.getText());
            int numNozzle2 = Integer.parseInt(textFieldNozzle2Quant.getText());
            GPM = (((GPM1 * numNozzle1) + (GPM2 * numNozzle2))/(numNozzle1 + numNozzle2))*Math.sqrt(PSI/40)*(numNozzle1+numNozzle2);
            return round(GPM, 2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            try{
                FileInputStream fis = new FileInputStream("NozzleModels.xlsx");
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sh = wb.getSheet(comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString());
                if(sh==null){
                    return -9;
                }
                XSSFRow row50 = sh.getRow(50);
                Iterator cells = row50.cellIterator();
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
                int numNozzle1 = Integer.parseInt(textFieldNozzle1Quant.getText());
                GPM = (sh.getRow(51).getCell(orifPos).getNumericCellValue())*Math.sqrt(PSI/40)*(numNozzle1);
                fis.close();
            } catch (Exception e){
                e.printStackTrace();
            }
            return round(GPM, 2);
        } else {
            return -9;
        }
    }

    public double calcEstGPA (){
        double GPA = 0;
        double GPM = calcEstGPM();
        if(GPM==-9){
            return -9;
        }
        double MPH=0;
        double numSpeed=0;
        if(!textFieldGS1.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS1.getText());numSpeed+=1;}
        if(!textFieldGS2.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS2.getText());numSpeed+=1;}
        if(!textFieldGS3.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS3.getText());numSpeed+=1;}
        if(!textFieldGS4.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS4.getText());numSpeed+=1;}
        if(!textFieldGS5.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS5.getText());numSpeed+=1;}
        if(!textFieldGS6.getText().isEmpty()){MPH += Double.parseDouble(textFieldGS6.getText());numSpeed+=1;}
        MPH = MPH/numSpeed;
        double SW = Double.parseDouble(labelSwathFinal.getText());
        GPA = (495*GPM)/(MPH*SW);
        return round(GPA, 2);
    }

    public double calcEstLPM (){
        double estGPM = calcEstGPM();
        if (estGPM!=-9) {
            return round(calcEstGPM()*3.78541, 2);
        } else {
            return -9;
        }
    }

    public double calcEstLPHA (){
        double GPA = calcEstGPA();
        if(GPA==-9){
            return -9;
        }
        double LPHA = GPA*3.78541*2.47105;
        return round(LPHA, 2);
    }

    public int getDV01(){
        int dV01 = 0;
        int dV01x = 0;
        int dV01xx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV01x = (int) Math.round(ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle));
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                dV01xx = (int) Math.round(ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle));
            }
            if(dV01x==-9 || dV01xx==-9){
                return -9;
            }
            dV01 = (dV01x*numNozzles1 + dV01xx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV01 = (int) Math.round(ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle));
            }
        } else {
            return -9;
        }
        return dV01;
    }

    public int getDV05(){
        int dV05 = 0;
        int dV05x = 0;
        int dV05xx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV05x = (int) Math.round(ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle));
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                dV05xx = (int) Math.round(ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle));
            }
            if(dV05x==-9 || dV05xx==-9){
                return -9;
            }
            dV05 = (dV05x*numNozzles1 + dV05xx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV05 = (int) Math.round(ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle));
            }
        } else {
            return -9;
        }
        return dV05;
    }

    public int getDV09(){
        int dV09 = 0;
        int dV09x = 0;
        int dV09xx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV09x = (int) Math.round(ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle));
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                dV09xx = (int) Math.round(ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle));
            }
            if(dV09x==-9 || dV09xx==-9){
                return -9;
            }
            dV09 = (dV09x*numNozzles1 + dV09xx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV09 = (int) Math.round(ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle));
            }
        } else {
            return -9;
        }
        return dV09;
    }

    public double getRS(){
        double rS = 0;
        double rSx = 0;
        double rSxx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                double dV01 = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                double dV05 = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);
                double dV09 = ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle);
                rSx = (dV09 - dV01) / dV05;
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                double dV01 = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                double dV05 = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);
                double dV09 = ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle);
                rSxx = (dV09 - dV01) / dV05;
            }
            if(rSx==-9 || rSxx==-9){
                return -9;
            }
            rS = (rSx*numNozzles1 + rSxx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                double dV01 = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                double dV05 = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);
                double dV09 = ModelCalc.getDV09(nozzle, orifice, airspeed, pressure, angle);
                rS = (dV09 - dV01) / dV05;
            }
        } else {
            return -9;
        }
        return round(rS, 2);
    }

    public double getPercentLess100(){
        double pL1 = 0;
        double pL1x = 0;
        double pL1xx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                pL1x = ModelCalc.getPercentLess100(nozzle, orifice, airspeed, pressure, angle);
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                pL1xx = ModelCalc.getPercentLess100(nozzle, orifice, airspeed, pressure, angle);
            }
            if(pL1x==-9 || pL1xx==-9){
                return -9;
            }
            pL1 = (pL1x*numNozzles1 + pL1xx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                pL1 = ModelCalc.getPercentLess100(nozzle, orifice, airspeed, pressure, angle);
            }
        } else {
            return -9;
        }
        if(pL1<0){
            pL1 = 0.01;
        }
        return pL1;
    }

    public String getDSC(){
        String dSC = "-";

        double dV01 = 0;
        double dV01x = 0;
        double dV01xx = 0;
        double dV05 = 0;
        double dV05x = 0;
        double dV05xx = 0;
        int numNozzles1 = Integer.parseInt(textFieldNozzle1Quant.getText());
        int numNozzles2 = 0;
        if(!textFieldNozzle2Quant.getText().isEmpty()){
            numNozzles2 = Integer.parseInt(textFieldNozzle2Quant.getText());
        }
        if(!textFieldNozzle1Quant.getText().isEmpty() && !textFieldNozzle2Quant.getText().isEmpty()){
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV01x = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                dV05x = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);
            }
            if (comboBoxNozzle2Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle2Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle2Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle2Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle2Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle2Def.getSelectionModel().getSelectedItem().toString());
                dV01xx = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                dV05xx = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);

            }
            if(dV01x==-9 || dV01xx==-9 || dV05x==-9 || dV05xx==-9){
                return "Unknown";
            }
            dV01 = (dV01x*numNozzles1 + dV01xx*numNozzles2)/(numNozzles1+numNozzles2);
            dV05 = (dV05x*numNozzles1 + dV05xx*numNozzles2)/(numNozzles1+numNozzles2);
        } else if (!textFieldNozzle1Quant.getText().isEmpty()) {
            if (comboBoxNozzle1Type.getSelectionModel().getSelectedItem() != "Select Nozzle" &&
                    comboBoxNozzle1Size.getSelectionModel().getSelectedItem() != "N/A" &&
                    !textFieldBoomPressure.getText().isEmpty() &&
                    comboBoxNozzle1Def.getSelectionModel().getSelectedItem() != "N/A") {
                String nozzle = comboBoxNozzle1Type.getSelectionModel().getSelectedItem().toString();
                double orifice = Double.parseDouble(comboBoxNozzle1Size.getSelectionModel().getSelectedItem().toString());
                double airspeed = averageD(airspeedCalc(1), airspeedCalc(2), airspeedCalc(3), airspeedCalc(4), airspeedCalc(5), airspeedCalc(6),2);
                double pressure = Double.parseDouble(textFieldBoomPressure.getText());
                double angle = Double.parseDouble(comboBoxNozzle1Def.getSelectionModel().getSelectedItem().toString());
                dV01 = ModelCalc.getDV01(nozzle, orifice, airspeed, pressure, angle);
                dV05 = ModelCalc.getDV05(nozzle, orifice, airspeed, pressure, angle);
                if(dV01==-9 || dV05==-9){
                    return "Unknown";
                }
            }
        } else {
            return "Unknown";
        }
        dSC = ModelCalc.getDSC(dV01, dV05);
        return dSC;
    }

    //External Class Methods
    static void setFileToOpen(File fileToOpen_){
        fileToOpen = fileToOpen_;
    }

    private File getCurrentDirectory(){
        return this.currentDirectory;
    }

    private void setCurrentDirectory(File directory){
        this.currentDirectory = directory;
    }

    private File getReportFile(){
        return this.reportFile;
    }

    public static void setHeaders(String headerOne, String headerTwo, String headerThree, String headerFour){
        header1 = headerOne;
        header2 = headerTwo;
        header3 = headerThree;
        header4 = headerFour;
    }

    public static void setMetric(boolean metric){
        METRIC = metric;
        if(METRIC){
            uAirSpeed = "MPH";
            uWindSpeed = "KPH";
            uPressure = "PSI";
            uTemp = "Deg C";
            uAppRate = "L/ha";
            uFlowRate = "L/min";
            uHeight = "m";
            uHeightSmall = "cm";
        } else {
            uAirSpeed = "MPH";
            uWindSpeed = "MPH";
            uPressure = "PSI";
            uTemp = "Deg F";
            uAppRate = "GPA";
            uFlowRate = "GPM";
            uHeight = "FT";
            uHeightSmall = "IN";
        }
    }

    public static void setFlightlineLength(double flightlineLength){
        FLIGHTLINE_LENGTH = flightlineLength;
    }

    public static void setStringVelocity(double stringVel){
        STRING_VEL = stringVel;
        STRING_TIME = FLIGHTLINE_LENGTH/STRING_VEL;
    }

    public static void setSampleLength(double sampleLength){
        SAMPLE_LENGTH  = sampleLength;
    }

    public static void setExcitationWavelength(int wavelength){
        EXCITATION_WAVELENGTH = wavelength;
    }

    public static void setTargetWavelength(int wavelength){
        TARGET_WAVELENGTH = wavelength;
    }

    public static void setBoxcarWidth(int width){
        BOXCAR_WIDTH = width;
        if(wrapper==null){
            return;
        }
        if (wrapper.getNumberOfSpectrometersFound()!=0) {
            wrapper.setBoxcarWidth(0,(width-1)/2);
        }
    }

    public static void setIntegrationTime(int integrationTime){
        INTEGRATION_TIME = integrationTime;
        if(wrapper==null){
            return;
        }
        if (wrapper.getNumberOfSpectrometersFound()!=0) {
            wrapper.setIntegrationTime(0, INTEGRATION_TIME);
        }
    }

    public static void setFlipPattern(boolean flipPattern){
        isFlipPattern = flipPattern;
    }

    public static void setUseLogo(boolean isUseLogo){
        useLogo = isUseLogo;
    }

    public static void setLogo(String newLogoFile){
        logoFile  = newLogoFile;
    }

    public static void setAutoInvertPH(boolean autoInvertPH){
        invertPH = autoInvertPH;
    }

    public static void setSpreadFactor(int factorIdentifier){
        useSpreadFactor = factorIdentifier;
    }

    public static void setDefaultThreshold(int thresh) {
        DEFAULT_THRESHOLD = thresh;}

    public static void setSpreadABC(double sprA, double sprB, double sprC){
        spreadA = sprA;
        spreadB = sprB;
        spreadC = sprC;
    }

    public static void setCardSpacing(int spacing){
        CARD_SPACING = spacing;
    }

    public static void setCropScale(double scale){
        CROP_SCALE = scale;
    }

}
