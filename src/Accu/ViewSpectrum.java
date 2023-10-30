package Accu;

import com.oceanoptics.omnidriver.api.wrapper.Wrapper;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.chart.AreaChart;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.util.Duration;

/**
 * Created by gill14 on 5/16/2016.
 */
public class ViewSpectrum {

    private AreaChart<Number, Number> areaChart;
    private XYChart.Series series;
    private XYChart.Series tPSeries;
    private double [] spectrum;
    private double [] wavelengths;
    private Timeline animation;

    private Wrapper wrapper;
    private int integrationTime;
    private int boxcarWidth;
    private int targetWavelength;
    private int excitationWavelength;

    private int targetPixel;
    private int excitationPixel;

    private static int boundLX = 400;
    private static int boundLXp;
    private static int boundRX = 800;
    private static int boundRXp;
    private static int boundY = 8000;

    ViewSpectrum(int integrationTime, int boxcarWidth, int excitationWavelength, int targetWavelength){
        this.excitationWavelength = excitationWavelength;
        this.integrationTime = integrationTime;
        this.boxcarWidth = boxcarWidth;
        this.targetWavelength = targetWavelength;
        if (testSpectrometer()) {
            setupSpectrometer();
            display();
        }
    }

    private boolean testSpectrometer(){
        wrapper = new Wrapper();
        wrapper.closeAllSpectrometers();
        int num = wrapper.openAllSpectrometers();
        if(num==1){
            return true;
        } else if (wrapper.openAllSpectrometers()>1){
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Spectrometer Connect Error");
            alert.setHeaderText("Multiple Instances Found");
            alert.setContentText("Close all versions of AccuPatt and try again.");
            alert.showAndWait();
            return false;
        } else {
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Spectrometer Connect Error");
            alert.setHeaderText("No Spectrometer found");
            alert.setContentText("Ensure Spectrometer is connected. If connected, check Windows Device Manager and ensure the driver \"Ocean Optics USB2000+ (WinUSB)\" is installed");
            alert.showAndWait();
            return false;
        }
    }

    private void setupSpectrometer(){

        wrapper.setCorrectForElectricalDark(0, 0);
        wrapper.setIntegrationTime(0,integrationTime/2);
        wavelengths = wrapper.getWavelengths(0);
        for(int i=0; i<wavelengths.length;i++){
            if (Math.round(wavelengths[i])==targetWavelength){
                targetPixel = i;
            }
            if (Math.round(wavelengths[i])==excitationWavelength){
                excitationPixel = i;
            }
            if (Math.round(wavelengths[i])==boundLX){
                boundLXp = i;
            }
            if (Math.round(wavelengths[i])==boundRX){
                boundRXp = i;
            }
        }
    }

    private void display() {
        Stage window = new Stage();
        window.initModality(Modality.APPLICATION_MODAL);
        window.setTitle("View Spectrum");
        window.setMinWidth(850);


        //Set Up AreaChart
        NumberAxis axisX = new NumberAxis();
        axisX.setAutoRanging(false);
        axisX.setLowerBound(boundLX);
        axisX.setUpperBound(boundRX);
        axisX.setLabel("Wavelength (nm)");
        NumberAxis axisY = new NumberAxis();
        axisY.setAutoRanging(false);
        axisY.setForceZeroInRange(false);
        axisY.setUpperBound(boundY);
        axisY.setLabel("Spectral Intensity");
        areaChart = new AreaChart<>(axisX, axisY);
        areaChart.setCreateSymbols(false);
        areaChart.setAnimated(false);
        areaChart.setLegendVisible(false);

        tPSeries = new AreaChart.Series();
        for(int q=boundY; q>0; q--){
            tPSeries.getData().add(new AreaChart.Data<>(wrapper.getWavelength(0,targetPixel-((boxcarWidth-1)/2)),q));
        }
        for(int q=0; q<boundY; q++){
            tPSeries.getData().add(new AreaChart.Data<>(wrapper.getWavelength(0,targetPixel),q));

        }
        tPSeries.getData().add(new AreaChart.Data<>(wrapper.getWavelength(0,targetPixel),0));
        for(int q=0; q<boundY; q++){
            tPSeries.getData().add(new AreaChart.Data<>(wrapper.getWavelength(0,targetPixel+((boxcarWidth-1)/2)),q));
        }
        areaChart.getData().addAll(tPSeries);

        //Sampling
        animation = new Timeline();
        animation.getKeyFrames().add(new KeyFrame(Duration.seconds(0.25), (event -> {
            series.getData().clear();
            spectrum = wrapper.getSpectrum(0);

            for(int i=boundLXp; i<boundRXp; i++) {

                series.getData().add(new AreaChart.Data<>(wavelengths[i], spectrum[i]));
            }
            })));
        animation.setCycleCount(5000);
        series = new AreaChart.Series();
        areaChart.getData().addAll(series);
        areaChart.setId("trimmedPass-Plot");
        animation.play();

        HBox hB0 = new HBox();
        hB0.setStyle("-fx-background-color: gray");
        Label label0 = new Label("Emission Wavelength: "+targetWavelength);
        Region reg02 = new Region();
        hB0.setHgrow(reg02, Priority.ALWAYS);
        Label label01 = new Label("Emission Pixel: "+targetPixel);
        Region reg03 = new Region();
        hB0.setHgrow(reg03, Priority.ALWAYS);
        Label label1 = new Label("Excitation Wavelength: "+excitationWavelength);
        Region reg04 = new Region();
        hB0.setHgrow(reg04, Priority.ALWAYS);
        Label label11 = new Label("Excitation Pixel: "+excitationPixel);
        hB0.getChildren().addAll(label0,reg02,label01, reg03, label1, reg04, label11);

        Label labelHint0 = new Label("You should see a band of intensity around your excitation wavelength if your LED is operating");
        Label labelHint01 = new Label("The vertical lines represent the averaging window (Boxcar Width) around the sampling (Emission) wavelength");

        VBox layout = new VBox();
        layout.getChildren().addAll(areaChart, labelHint0, labelHint01, hB0);
        layout.setAlignment(Pos.CENTER);

        Scene scene = new Scene(layout);
        scene.getStylesheets().add(Main.class.getResource("Accu.css").toExternalForm());
        window.setScene(scene);
        window.setOnCloseRequest(e -> animation.stop());
        window.showAndWait();


    }
}
