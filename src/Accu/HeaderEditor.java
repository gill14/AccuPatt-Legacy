package Accu;

import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;

/**
 * Created by gill14 on 1/7/2016.
 */
public class HeaderEditor {

    public static void display(String header1, String header2, String header3, String header4){
        Stage window = new Stage();
        window.initModality(Modality.APPLICATION_MODAL);
        window.setTitle("Edit Report Headers");
        window.setMinWidth(350);
        window.setMinHeight(250);

        TextField textFieldHeader1 = new TextField();
        if (header1 != null) {
            textFieldHeader1.setText(header1);
        } else {
            textFieldHeader1.setPromptText("Fly-In");
        }
        TextField textFieldHeader2 = new TextField();
        if (header2 != null) {
            textFieldHeader2.setText(header2);
        } else {
            textFieldHeader2.setPromptText("Location");
        }
        TextField textFieldHeader3 = new TextField();
        if (header3 != null) {
            textFieldHeader3.setText(header3);
        } else {
            textFieldHeader3.setPromptText("Date");
        }
        TextField textFieldHeader4 = new TextField();
        if (header4 != null) {
            textFieldHeader4.setText(header4);
        } else {
            textFieldHeader4.setPromptText("Analyst");
        }

        Button saveIt = new Button("Save");
        saveIt.setOnAction(e -> {
            Controller.setHeaders(textFieldHeader1.getText(),
                    textFieldHeader2.getText(),
                    textFieldHeader3.getText(),
                    textFieldHeader4.getText());
            window.close();
        });


        VBox layout = new VBox(10);
        layout.getChildren().addAll(textFieldHeader1, textFieldHeader2, textFieldHeader3, textFieldHeader4, saveIt);
        layout.setAlignment(Pos.CENTER);

        Scene scene = new Scene(layout);
        window.setScene(scene);
        window.showAndWait();
    }

}
