package Accu;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;

/**
 * Created by gill14 on 11/19/18.
 */
public class LineEntry {
    private final StringProperty header1;
    private final StringProperty header2;
    private final StringProperty header3;
    private final StringProperty header4;
    //
    private final StringProperty pilot;
    private final StringProperty bus;
    private final StringProperty state;
    private final StringProperty regnum;
    private final StringProperty series;
    private final StringProperty make;
    private final StringProperty model;
    private final StringProperty psi;
    private final StringProperty gpa;
    private final StringProperty noz1T;
    private final StringProperty noz1Q;
    private final StringProperty noz1O;
    private final StringProperty noz1D;
    private final StringProperty noz2T;
    private final StringProperty noz2Q;
    private final StringProperty noz2O;
    private final StringProperty noz2D;
    private final StringProperty tsw;
    private final StringProperty psw;
    private final StringProperty pbw;
    private final StringProperty ns;
    private final StringProperty winglets;
    private final StringProperty notes;
    private final StringProperty file;

    public LineEntry(){
        this(null,null,null,null,null,null,null,null,null,null,
                null,null,null,null,null,null,null,null,null,null,
                null,null,null,null,null,null,null, null);

    }

    public LineEntry(String header1, String header2, String header3, String header4, String pilot, String bus, String state, String regnum,
                     String series, String make, String model, String psi, String gpa, String noz1T, String noz1Q, String noz1O, String noz1D,
                     String noz2T, String noz2Q, String noz2O, String noz2D, String tsw, String psw, String pbw, String ns, String winglets,
                     String notes, String file){
        this.header1 = new SimpleStringProperty(header1);
        this.header2 = new SimpleStringProperty(header2);
        this.header3 = new SimpleStringProperty(header3);
        this.header4 = new SimpleStringProperty(header4);
        this.pilot = new SimpleStringProperty(pilot);
        this.bus = new SimpleStringProperty(bus);
        this.state = new SimpleStringProperty(state);
        this.regnum = new SimpleStringProperty(regnum);
        this.series = new SimpleStringProperty(series);
        this.make = new SimpleStringProperty(make);
        this.model = new SimpleStringProperty(model);
        this.psi = new SimpleStringProperty(psi);
        this.gpa = new SimpleStringProperty(gpa);
        this.noz1T = new SimpleStringProperty(noz1T);
        this.noz1Q = new SimpleStringProperty(noz1Q);
        this.noz1O = new SimpleStringProperty(noz1O);
        this.noz1D = new SimpleStringProperty(noz1D);
        this.noz2T = new SimpleStringProperty(noz2T);
        this.noz2Q = new SimpleStringProperty(noz2Q);
        this.noz2O = new SimpleStringProperty(noz2O);
        this.noz2D = new SimpleStringProperty(noz2D);
        this.tsw = new SimpleStringProperty(tsw);
        this.psw = new SimpleStringProperty(psw);
        this.pbw = new SimpleStringProperty(pbw);
        this.ns = new SimpleStringProperty(ns);
        this.winglets = new SimpleStringProperty(winglets);
        this.notes = new SimpleStringProperty(notes);
        this.file = new SimpleStringProperty(file);
    }

    public String getheader1() {
        return header1.get();
    }

    public StringProperty header1Property() {
        return header1;
    }

    public String getheder2() {
        return header2.get();
    }

    public StringProperty header2Property() {
        return header2;
    }

    public String getheader3() {
        return header3.get();
    }

    public StringProperty header3Property() {
        return header3;
    }

    public String getheder4() {
        return header4.get();
    }

    public StringProperty header4Property() {
        return header4;
    }

    public String getPilot() {
        return pilot.get();
    }

    public StringProperty pilotProperty() {
        return pilot;
    }

    public String getBus() {
        return bus.get();
    }

    public StringProperty busProperty() {
        return bus;
    }

    public String getState() {
        return state.get();
    }

    public StringProperty stateProperty() {
        return state;
    }

    public String getRegnum() {
        return regnum.get();
    }

    public StringProperty regnumProperty() {
        return regnum;
    }

    public String getSeries() { return series.get();}

    public StringProperty seriesProperty() {return series;}

    public String getMake() {
        return make.get();
    }

    public StringProperty makeProperty() {
        return make;
    }

    public String getModel() {
        return model.get();
    }

    public StringProperty modelProperty() {
        return model;
    }

    public String getPsi() {
        return psi.get();
    }

    public StringProperty psiProperty() {
        return psi;
    }

    public String getGpa() {
        return gpa.get();
    }

    public StringProperty gpaProperty() {
        return gpa;
    }

    public String getNoz1T() {
        return noz1T.get();
    }

    public StringProperty noz1TProperty() {
        return noz1T;
    }

    public String getNoz1Q() {
        return noz1Q.get();
    }

    public StringProperty noz1QProperty() {
        return noz1Q;
    }

    public String getNoz1O() {
        return noz1O.get();
    }

    public StringProperty noz1OProperty() {
        return noz1O;
    }

    public String getNoz1D() {
        return noz1D.get();
    }

    public StringProperty noz1DProperty() {
        return noz1D;
    }

    public String getNoz2T() {
        return noz2T.get();
    }

    public StringProperty noz2TProperty() {
        return noz2T;
    }

    public String getNoz2Q() {
        return noz2Q.get();
    }

    public StringProperty noz2QProperty() {
        return noz2Q;
    }

    public String getNoz2O() {
        return noz2O.get();
    }

    public StringProperty noz2OProperty() {
        return noz2O;
    }

    public String getNoz2D() {
        return noz2D.get();
    }

    public StringProperty noz2DProperty() {
        return noz2D;
    }

    public String getTsw() {
        return tsw.get();
    }

    public StringProperty tswProperty() {
        return tsw;
    }

    public String getPsw() {
        return psw.get();
    }

    public StringProperty pswProperty() {
        return psw;
    }

    public String getPbw() {
        return pbw.get();
    }

    public StringProperty pbwProperty() {
        return pbw;
    }

    public String getNs() {
        return ns.get();
    }

    public StringProperty nsProperty() {
        return ns;
    }

    public String getWinglets() {
        return winglets.get();
    }

    public StringProperty wingletsProperty() {
        return winglets;
    }

    public String getNotes() {
        return notes.get();
    }

    public StringProperty notesProperty() {
        return notes;
    }

    public String getFile() {
        return file.get();
    }

    public StringProperty fileProperty() {
        return file;
    }
}
