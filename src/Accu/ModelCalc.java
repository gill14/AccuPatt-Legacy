package Accu;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by gill14 on 2/3/2016.
 */
public class ModelCalc {

    //Nozzle Model File
    public static String nozzleModelFile = "NozzleModels.xlsx";

    //Reference Nozzles
    public static double dV01VFF = 59.5;
    public static double dV01FM = 110.3;
    public static double dV01MC = 162;
    public static double dV01CVC = 191.7;
    public static double dV01VCXC = 226.1;
    public static double dV05VFF = 134.4;
    public static double dV05FM = 248.1;
    public static double dV05MC = 357.8;
    public static double dV05CVC = 431;
    public static double dV05VCXC = 500.9;
    public static double dV09VFF = 236.4;
    public static double dV09FM = 409.9;
    public static double dV09MC = 584;
    public static double dV09CVC = 737.1;
    public static double dV09VCXC = 819.8;


    public static double getDV01 (String nozzle, double orifice, double airspeed, double pressure, double angle){
        //CCD Factors
        double orfSub = 0;
        double orfDiv = 0;
        double airSub = 0;
        double airDiv = 0;
        double preSub = 0;
        double preDiv = 0;
        double angSub = 0;
        double angDiv = 0;
        //Terms
        double [] terms = new double [15];
        try {
            FileInputStream fis = new FileInputStream(nozzleModelFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet(nozzle);
            int CCDRow;
            int VRow;
            if (sh != null) {
                if(airspeed >= sh.getRow(1).getCell(1).getNumericCellValue() && airspeed <= sh.getRow(1).getCell(2).getNumericCellValue()){
                    //Use High Speed Model
                    CCDRow = 7;
                    VRow = 10;
                } else if (airspeed < sh.getRow(1).getCell(1).getNumericCellValue()){
                    //Use Low Speed Model
                    if(airspeed >= sh.getRow(16).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(16).getCell(2).getNumericCellValue() &&
                            angle >= sh.getRow(19).getCell(1).getNumericCellValue()){
                        //Use universal or def
                        CCDRow = 22;
                        VRow = 25;
                    } else if (airspeed >= sh.getRow(31).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(31).getCell(2).getNumericCellValue()){
                        //Use alternate Low Speed Model
                        CCDRow = 37;
                        VRow = 40;
                    } else {
                        //Too Slow, Outside Params
                        return -9;
                    }
                } else {
                    //Too Fast, Outside Params
                    return -9;
                }
            } else {
                //No model data for named nozzle
                return -9;
            }

            orfSub = sh.getRow(CCDRow).getCell(1).getNumericCellValue();
            orfDiv = sh.getRow(CCDRow).getCell(2).getNumericCellValue();
            airSub = sh.getRow(CCDRow).getCell(3).getNumericCellValue();
            airDiv = sh.getRow(CCDRow).getCell(4).getNumericCellValue();
            preSub = sh.getRow(CCDRow).getCell(5).getNumericCellValue();
            preDiv = sh.getRow(CCDRow).getCell(6).getNumericCellValue();
            angSub = sh.getRow(CCDRow).getCell(7).getNumericCellValue();
            angDiv = sh.getRow(CCDRow).getCell(8).getNumericCellValue();
            for(int i=0; i<15; i++){
                terms[i] = sh.getRow(VRow).getCell(i+1).getNumericCellValue();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        double CCDOrf = (orifice - orfSub)/orfDiv;
        double CCDAir = (airspeed - airSub)/airDiv;
        double CCDPre = (pressure - preSub)/preDiv;
        double CCDAng = (angle - angSub)/angDiv;

        double dV01 = terms[0] + (CCDOrf * terms[1]) + (CCDAir * terms[2]) + (CCDPre * terms[3]) +
                (CCDAng * terms[4]) + (CCDOrf * CCDAir * terms[5]) + (CCDOrf * CCDPre * terms[6]) +
                (CCDAir * CCDPre * terms[7]) + (CCDOrf * CCDAng * terms[8])  + (CCDAir * CCDAng * terms[9]) +
                (CCDPre * CCDAng * terms[10]) + (CCDOrf * CCDOrf * terms[11]) + (CCDAir * CCDAir * terms[12]) +
                (CCDPre * CCDPre * terms[13]) + (CCDAng * CCDAng * terms[14]);

        return dV01;

    }

    public static double getDV05 (String nozzle, double orifice, double airspeed, double pressure, double angle){
        //CCD Factors
        double orfSub = 0;
        double orfDiv = 0;
        double airSub = 0;
        double airDiv = 0;
        double preSub = 0;
        double preDiv = 0;
        double angSub = 0;
        double angDiv = 0;
        //Terms
        double [] terms = new double [15];
        try {
            FileInputStream fis = new FileInputStream(nozzleModelFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet(nozzle);
            int CCDRow;
            int VRow;
            if (sh != null) {
                if(airspeed >= sh.getRow(1).getCell(1).getNumericCellValue() && airspeed <= sh.getRow(1).getCell(2).getNumericCellValue()){
                    //Use High Speed Model
                    CCDRow = 7;
                    VRow = 11;
                } else if (airspeed < sh.getRow(1).getCell(1).getNumericCellValue()){
                    //Use Low Speed Model
                    if(airspeed >= sh.getRow(16).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(16).getCell(2).getNumericCellValue() &&
                            angle >= sh.getRow(19).getCell(1).getNumericCellValue()){
                        //Use universal or def
                        CCDRow = 22;
                        VRow = 26;
                    } else if (airspeed >= sh.getRow(31).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(31).getCell(2).getNumericCellValue()){
                        //Use alternate Low Speed Model
                        CCDRow = 37;
                        VRow = 41;
                    } else {
                        //Too Slow, Outside Params
                        return -9;
                    }
                } else {
                    //Too Fast, Outside Params
                    return -9;
                }
            } else {
                //No model data for named nozzle
                return -9;
            }

            orfSub = sh.getRow(CCDRow).getCell(1).getNumericCellValue();
            orfDiv = sh.getRow(CCDRow).getCell(2).getNumericCellValue();
            airSub = sh.getRow(CCDRow).getCell(3).getNumericCellValue();
            airDiv = sh.getRow(CCDRow).getCell(4).getNumericCellValue();
            preSub = sh.getRow(CCDRow).getCell(5).getNumericCellValue();
            preDiv = sh.getRow(CCDRow).getCell(6).getNumericCellValue();
            angSub = sh.getRow(CCDRow).getCell(7).getNumericCellValue();
            angDiv = sh.getRow(CCDRow).getCell(8).getNumericCellValue();
            for(int i=0; i<15; i++){
                terms[i] = sh.getRow(VRow).getCell(i+1).getNumericCellValue();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        double CCDOrf = (orifice - orfSub)/orfDiv;
        double CCDAir = (airspeed - airSub)/airDiv;
        double CCDPre = (pressure - preSub)/preDiv;
        double CCDAng = (angle - angSub)/angDiv;

        double dV05 = terms[0] + (CCDOrf * terms[1]) + (CCDAir * terms[2]) + (CCDPre * terms[3]) +
                (CCDAng * terms[4]) + (CCDOrf * CCDAir * terms[5]) + (CCDOrf * CCDPre * terms[6]) +
                (CCDAir * CCDPre * terms[7]) + (CCDOrf * CCDAng * terms[8])  + (CCDAir * CCDAng * terms[9]) +
                (CCDPre * CCDAng * terms[10]) + (CCDOrf * CCDOrf * terms[11]) + (CCDAir * CCDAir * terms[12]) +
                (CCDPre * CCDPre * terms[13]) + (CCDAng * CCDAng * terms[14]);

        return dV05;

    }

    public static double getDV09 (String nozzle, double orifice, double airspeed, double pressure, double angle){
        //CCD Factors
        double orfSub = 0;
        double orfDiv = 0;
        double airSub = 0;
        double airDiv = 0;
        double preSub = 0;
        double preDiv = 0;
        double angSub = 0;
        double angDiv = 0;
        //Terms
        double [] terms = new double [15];
        try {
            FileInputStream fis = new FileInputStream(nozzleModelFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet(nozzle);
            int CCDRow;
            int VRow;
            if (sh != null) {
                if(airspeed >= sh.getRow(1).getCell(1).getNumericCellValue() && airspeed <= sh.getRow(1).getCell(2).getNumericCellValue()){
                    //Use High Speed Model
                    CCDRow = 7;
                    VRow = 12;
                } else if (airspeed < sh.getRow(1).getCell(1).getNumericCellValue()){
                    //Use Low Speed Model
                    if(airspeed >= sh.getRow(16).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(16).getCell(2).getNumericCellValue() &&
                            angle >= sh.getRow(19).getCell(1).getNumericCellValue()){
                        //Use universal or def
                        CCDRow = 22;
                        VRow = 27;
                    } else if (airspeed >= sh.getRow(31).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(31).getCell(2).getNumericCellValue()){
                        //Use alternate Low Speed Model
                        CCDRow = 37;
                        VRow = 42;
                    } else {
                        //Too Slow, Outside Params
                        return -9;
                    }
                } else {
                    //Too Fast, Outside Params
                    return -9;
                }
            } else {
                //No model data for named nozzle
                return -9;
            }

            orfSub = sh.getRow(CCDRow).getCell(1).getNumericCellValue();
            orfDiv = sh.getRow(CCDRow).getCell(2).getNumericCellValue();
            airSub = sh.getRow(CCDRow).getCell(3).getNumericCellValue();
            airDiv = sh.getRow(CCDRow).getCell(4).getNumericCellValue();
            preSub = sh.getRow(CCDRow).getCell(5).getNumericCellValue();
            preDiv = sh.getRow(CCDRow).getCell(6).getNumericCellValue();
            angSub = sh.getRow(CCDRow).getCell(7).getNumericCellValue();
            angDiv = sh.getRow(CCDRow).getCell(8).getNumericCellValue();
            for(int i=0; i<15; i++){
                terms[i] = sh.getRow(VRow).getCell(i+1).getNumericCellValue();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        double CCDOrf = (orifice - orfSub)/orfDiv;
        double CCDAir = (airspeed - airSub)/airDiv;
        double CCDPre = (pressure - preSub)/preDiv;
        double CCDAng = (angle - angSub)/angDiv;

        double dV09 = terms[0] + (CCDOrf * terms[1]) + (CCDAir * terms[2]) + (CCDPre * terms[3]) +
                (CCDAng * terms[4]) + (CCDOrf * CCDAir * terms[5]) + (CCDOrf * CCDPre * terms[6]) +
                (CCDAir * CCDPre * terms[7]) + (CCDOrf * CCDAng * terms[8])  + (CCDAir * CCDAng * terms[9]) +
                (CCDPre * CCDAng * terms[10]) + (CCDOrf * CCDOrf * terms[11]) + (CCDAir * CCDAir * terms[12]) +
                (CCDPre * CCDPre * terms[13]) + (CCDAng * CCDAng * terms[14]);

        return dV09;

    }

    public static double getPercentLess100 (String nozzle, double orifice, double airspeed, double pressure, double angle){
        //CCD Factors
        double orfSub = 0;
        double orfDiv = 0;
        double airSub = 0;
        double airDiv = 0;
        double preSub = 0;
        double preDiv = 0;
        double angSub = 0;
        double angDiv = 0;
        //Terms
        double [] terms = new double [15];
        try {
            FileInputStream fis = new FileInputStream(nozzleModelFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sh = wb.getSheet(nozzle);
            int CCDRow;
            int VRow;
            if (sh != null) {
                if(airspeed >= sh.getRow(1).getCell(1).getNumericCellValue() && airspeed <= sh.getRow(1).getCell(2).getNumericCellValue()){
                    //Use High Speed Model
                    CCDRow = 7;
                    VRow = 13;
                } else if (airspeed < sh.getRow(1).getCell(1).getNumericCellValue()){
                    //Use Low Speed Model
                    if(airspeed >= sh.getRow(16).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(16).getCell(2).getNumericCellValue() &&
                            angle >= sh.getRow(19).getCell(1).getNumericCellValue()){
                        //Use universal or def
                        CCDRow = 22;
                        VRow = 28;
                    } else if (airspeed >= sh.getRow(31).getCell(1).getNumericCellValue() &&
                            airspeed <= sh.getRow(31).getCell(2).getNumericCellValue()){
                        //Use alternate Low Speed Model
                        CCDRow = 37;
                        VRow = 43;
                    } else {
                        //Too Slow, Outside Params
                        return -9;
                    }
                } else {
                    //Too Fast, Outside Params
                    return -9;
                }
            } else {
                //No model data for named nozzle
                return -9;
            }

            orfSub = sh.getRow(CCDRow).getCell(1).getNumericCellValue();
            orfDiv = sh.getRow(CCDRow).getCell(2).getNumericCellValue();
            airSub = sh.getRow(CCDRow).getCell(3).getNumericCellValue();
            airDiv = sh.getRow(CCDRow).getCell(4).getNumericCellValue();
            preSub = sh.getRow(CCDRow).getCell(5).getNumericCellValue();
            preDiv = sh.getRow(CCDRow).getCell(6).getNumericCellValue();
            angSub = sh.getRow(CCDRow).getCell(7).getNumericCellValue();
            angDiv = sh.getRow(CCDRow).getCell(8).getNumericCellValue();
            for(int i=0; i<15; i++){
                terms[i] = sh.getRow(VRow).getCell(i+1).getNumericCellValue();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        double CCDOrf = (orifice - orfSub)/orfDiv;
        double CCDAir = (airspeed - airSub)/airDiv;
        double CCDPre = (pressure - preSub)/preDiv;
        double CCDAng = (angle - angSub)/angDiv;

        double percentLess100 = terms[0] + (CCDOrf * terms[1]) + (CCDAir * terms[2]) + (CCDPre * terms[3]) +
                (CCDAng * terms[4]) + (CCDOrf * CCDAir * terms[5]) + (CCDOrf * CCDPre * terms[6]) +
                (CCDAir * CCDPre * terms[7]) + (CCDOrf * CCDAng * terms[8])  + (CCDAir * CCDAng * terms[9]) +
                (CCDPre * CCDAng * terms[10]) + (CCDOrf * CCDOrf * terms[11]) + (CCDAir * CCDAir * terms[12]) +
                (CCDPre * CCDPre * terms[13]) + (CCDAng * CCDAng * terms[14]);

        return percentLess100;

    }

    public static String getDSC (double dV01, double dV05){
        //Check if Models exist first
        if(dV01==-9 || dV05==-9){
            return "N/A";
        }
        double dV01Num;
        if(dV01 < dV01VCXC){
            if(dV01 < dV01CVC){
                if(dV01 < dV01MC){
                    if(dV01 <  dV01FM){
                        if(dV01 < dV01VFF){
                            dV01Num = 1;
                        } else {
                            dV01Num = 2;
                        }
                    } else {
                        dV01Num = 3;
                    }
                } else {
                    dV01Num = 4;
                }
            } else {
                dV01Num = 5;
            }
        } else {
            dV01Num = 6;
        }
        double dV05Num;
        if(dV05 < dV05VCXC){
            if(dV05 < dV05CVC){
                if(dV05 < dV05MC){
                    if(dV05 < dV05FM){
                        if(dV05 < dV05VFF){
                            dV05Num = 1;
                        } else {
                            dV05Num = 2;
                        }
                    } else {
                        dV05Num = 3;
                    }
                } else {
                    dV05Num = 4;
                }
            } else {
                dV05Num = 5;
            }
        } else {
            dV05Num = 6;
        }

        double dSCNum;
        if(dV05Num < dV01Num){
            dSCNum = dV05Num;
        } else if (dV01Num < dV05Num){
            dSCNum = dV01Num;
        } else {
            dSCNum = dV05Num;
        }

        if(dSCNum == 1){
            return "VF";
        } else if(dSCNum == 2){
            return "F";
        } else if(dSCNum == 3) {
            return "M";
        } else if(dSCNum == 4){
            return "C";
        } else if(dSCNum == 5){
            return "VC";
        } else if(dSCNum == 6){
            return "XC";
        } else {
            return "";
        }

    }
}
