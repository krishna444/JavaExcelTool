package com.kpaudel;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.DatePicker;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

import java.time.LocalDate;

/**
 * @author krishna<br>
 * Date 28.10.19
 */
public class MainApp extends Application {
    @Override
    public void start(Stage primaryStage) throws Exception {
        primaryStage.setTitle("Tätigkeitsnachweise");
        StackPane root=new StackPane();

        DatePicker datePicker=new DatePicker(LocalDate.now());
        primaryStage.setScene(new Scene(root,300,250));
        primaryStage.show();
    }
}
