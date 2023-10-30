package Accu;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;


public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {
        setUserAgentStylesheet(STYLESHEET_CASPIAN);
        Parent root = FXMLLoader.load(getClass().getResource("AccuPatt_Home.fxml"));
        primaryStage.setTitle("AccuPatt - Aircraft Pattern Testing Analysis");
        Scene scene = new Scene(root, 800, 1000);
        primaryStage.setScene(scene);
        primaryStage.setMaximized(true);
        Image icon = new Image(getClass().getResourceAsStream("Resources/aLogo1.png"));
        primaryStage.getIcons().add(icon);
        scene.getStylesheets().add(Main.class.getResource("Accu.css").toExternalForm());
        primaryStage.show();

        primaryStage.setOnCloseRequest(e -> {
            Platform.exit();
            System.exit(0);
        });

    }


    public static void main(String[] args) {
        launch(args);
    }

}
