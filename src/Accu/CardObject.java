package Accu;

import ij.ImagePlus;
import ij.gui.Overlay;
import ij.plugin.frame.RoiManager;
import javafx.scene.image.Image;


import java.awt.image.BufferedImage;
import java.util.ArrayList;

/**
 * Created by gill14 on 4/5/19.
 */
public class CardObject {

    CardObject(int defaultThreshold){
        imp = new ImagePlus();
        rm = new RoiManager(false);
        overlay = new Overlay();

        setAllThresholds(defaultThreshold);

        cardList = new ArrayList<>();
        usingCardList = new ArrayList<>();
        cardNameList = new ArrayList<>();

        cardBufferedImages = new ArrayList<>();
        cardImages = new ArrayList<>();
        cardImagesThresholded = new ArrayList<>();

        dV01s = new ArrayList<>();
        dV05s = new ArrayList<>();
        dV09s = new ArrayList<>();
        dropsPerSquareInch = new ArrayList<>();
        percentCoverage = new ArrayList<>();
        cardAreas = new ArrayList<>();

        stains = new ArrayList<>();
        diameters = new ArrayList<>();
        volumes = new ArrayList<>();
    }

    private ImagePlus imp;
    private RoiManager rm;
    private Overlay overlay;

    ArrayList<Integer> thresholds;

    ArrayList<Boolean> cardList;
    ArrayList<Boolean> usingCardList;
    ArrayList<String> cardNameList;

    ArrayList<BufferedImage> cardBufferedImages;
    ArrayList<Image> cardImages;
    ArrayList<Image> cardImagesThresholded;

    ArrayList<ArrayList<Double>> stains;
    ArrayList<ArrayList<Double>> diameters;
    ArrayList<ArrayList<Double>> volumes;

    ArrayList<Double> dV01s;
    ArrayList<Double> dV05s;
    ArrayList<Double> dV09s;
    ArrayList<Double> dropsPerSquareInch;
    ArrayList<Double> percentCoverage;
    ArrayList<Double> cardAreas;



    //Useful getters/setters
    private void setAllThresholds(int thresholdValue){
        thresholds = new ArrayList<>();
        for(int i=0; i<9; i++){
            thresholds.add(thresholdValue);
        }
    }

    //variable getters/setters

    public ImagePlus getImp() {
        return imp;
    }

    public void setImp(ImagePlus imp) {
        this.imp = imp;
    }

    public RoiManager getRm() {
        return rm;
    }

    public void setRm(RoiManager rm) {
        this.rm = rm;
    }

    public Overlay getOverlay() {
        return overlay;
    }

    public void setOverlay(Overlay overlay) {
        this.overlay = overlay;
    }

    public ArrayList<Integer> getThresholds() {
        return thresholds;
    }

    public void setThresholds(ArrayList<Integer> thresholds) {
        this.thresholds = thresholds;
    }

    public ArrayList<Boolean> getCardList() {
        return cardList;
    }

    public void setCardList(ArrayList<Boolean> cardList) {
        this.cardList = cardList;
    }

    public ArrayList<Boolean> getUsingCardList() {
        return usingCardList;
    }

    public void setUsingCardList(ArrayList<Boolean> usingCardList) {
        this.usingCardList = usingCardList;
    }

    public ArrayList<String> getCardNameList() {
        return cardNameList;
    }

    public void setCardNameList(ArrayList<String> cardNameList) {
        this.cardNameList = cardNameList;
    }


    public ArrayList<Double> getdV01s() {
        return dV01s;
    }

    public void setdV01s(ArrayList<Double> dV01s) {
        this.dV01s = dV01s;
    }

    public ArrayList<Double> getdV05s() {
        return dV05s;
    }

    public void setdV05s(ArrayList<Double> dV05s) {
        this.dV05s = dV05s;
    }

    public ArrayList<Double> getdV09s() {
        return dV09s;
    }

    public void setdV09s(ArrayList<Double> dV09s) {
        this.dV09s = dV09s;
    }

    public ArrayList<Double> getDropsPerSquareInch() {
        return dropsPerSquareInch;
    }

    public void setDropsPerSquareInch(ArrayList<Double> dropsPerSquareInch) {
        this.dropsPerSquareInch = dropsPerSquareInch;
    }

    public ArrayList<Double> getPercentCoverage() {
        return percentCoverage;
    }

    public void setPercentCoverage(ArrayList<Double> percentCoverage) {
        this.percentCoverage = percentCoverage;
    }

    public ArrayList<BufferedImage> getCardBufferedImages() {
        return cardBufferedImages;
    }

    public void setCardBufferedImages(ArrayList<BufferedImage> cardBufferedImages) {
        this.cardBufferedImages = cardBufferedImages;
    }

    public ArrayList<Image> getCardImages() {
        return cardImages;
    }

    public void setCardImages(ArrayList<Image> cardImages) {
        this.cardImages = cardImages;
    }

    public ArrayList<Image> getCardImagesThresholded() {
        return cardImagesThresholded;
    }

    public void setCardImagesThresholded(ArrayList<Image> cardImagesThresholded) {
        this.cardImagesThresholded = cardImagesThresholded;
    }

    public ArrayList<ArrayList<Double>> getStains() {
        return stains;
    }

    public void setStains(ArrayList<ArrayList<Double>> stains) {
        this.stains = stains;
    }

    public ArrayList<ArrayList<Double>> getDiameters() {
        return diameters;
    }

    public void setDiameters(ArrayList<ArrayList<Double>> diameters) {
        this.diameters = diameters;
    }

    public ArrayList<ArrayList<Double>> getVolumes() {
        return volumes;
    }

    public void setVolumes(ArrayList<ArrayList<Double>> volumes) {
        this.volumes = volumes;
    }

    public ArrayList<Double> getCardAreas() {
        return cardAreas;
    }

    public void setCardAreas(ArrayList<Double> cardAreas) {
        this.cardAreas = cardAreas;
    }

}
