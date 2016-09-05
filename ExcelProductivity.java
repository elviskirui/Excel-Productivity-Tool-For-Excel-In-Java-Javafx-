
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.sql.*;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBuilder;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.scene.text.TextBuilder;
import javafx.scene.web.HTMLEditor;
import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProductivity extends Application {

    String db_host = "localhost";
    String db_user = "root";
    String db_pass = "kevo";
    String url = "jdbc:mysql://localhost:3306/";
    String db_name = "insurance";
    String jdbc = "com.mysql.jdbc.Driver";
    Text url_text;
    //connection strings
    PreparedStatement ps;
    Connection conn;
    ResultSet rs;

    public static void main(String[] args) {
        Application.launch(args);
    }

    public void start(Stage primaryStage) {
        
        //*************** BorderPane for export ************
        BorderPane export_db_detailsBP = new BorderPane();
        export_db_detailsBP.setId("pane");
        Scene scene_export_dbS = new Scene(export_db_detailsBP,600,400);
        scene_export_dbS.getStylesheets().add("ecss.css");
        
        //nodes for borderpane export
        
        TextField exportUrlTF = new TextField();
        exportUrlTF.setEditable(false);
        exportUrlTF.setPrefWidth(250);
        exportUrlTF.setPrefHeight(20);
        
       
        
        
        
        
        Button search_save_destinationB = new Button("Path");
        search_save_destinationB.setId("button");
        
        Button export_button = new Button("Export");
        export_button.setId("button");
        
        Button back_from_exportS = new Button("Back");
        back_from_exportS.setId("button");
        
        
        
        
        HBox export_hb = new HBox(10);
        export_hb.setAlignment(Pos.CENTER);
        export_hb.getChildren().addAll(exportUrlTF,export_button,back_from_exportS);
        export_db_detailsBP.setCenter(export_hb);
        
        
        
        export_button.setOnAction((ActionEvent t) -> {
            
            FileChooser excel_pathFC = new FileChooser();
            FileChooser.ExtensionFilter efc = new FileChooser.ExtensionFilter("excel files (*.xls)", "*.xls");
            excel_pathFC.getExtensionFilters().add(efc);
            String excel_file_S = excel_pathFC.showSaveDialog(primaryStage).toString();
            exportUrlTF.setText(excel_file_S);
            
            
            
            try {
                Class.forName(jdbc);
                conn = DriverManager.getConnection(url + db_name, db_user, db_pass);
                PreparedStatement psx = conn.prepareStatement("SELECT * FROM users");
                rs = psx.executeQuery();
                
                
                XSSFWorkbook wb = new XSSFWorkbook();
                XSSFSheet sheet = wb.createSheet("Client Details");
                XSSFRow titles = sheet.createRow(0);
                titles.createCell(0).setCellValue("ID");
                titles.createCell(1).setCellValue("First Name");
                titles.createCell(2).setCellValue("Last Name");
                titles.createCell(3).setCellValue("Email");
                
                
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                sheet.autoSizeColumn(3);
                sheet.autoSizeColumn(4);
                sheet.setColumnWidth(3,256 * 25);
                int index = 1;
                while(rs.next())
                {
                    XSSFRow row = sheet.createRow(index);
                    row.createCell(0).setCellValue(rs.getString("id"));
                    row.createCell(1).setCellValue(rs.getString("fname"));
                    row.createCell(2).setCellValue(rs.getString("lname"));
                    row.createCell(3).setCellValue(rs.getString("email"));
                    index ++;
                }
                
                
                
                FileOutputStream fileout = new FileOutputStream(excel_file_S);
                wb.write(fileout);
                fileout.close();
                
                
                Alert info = new Alert(AlertType.INFORMATION);
                info.setTitle("Information Dialog");
                info.setHeaderText("Information Dialog For Successful Excel File Created.");
                info.setContentText("Excel File Created Successfully!");
                info.showAndWait();
                
                
                
                
                psx.close();
                rs.close();
            } catch (ClassNotFoundException ex) {
                Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
            } catch (SQLException ex) {
                Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
            } catch (FileNotFoundException ex) {
                Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
            }
                    
            
        });
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        //************************ BorderPane for insurance *************
        BorderPane import_insuranceBP = new BorderPane();
        import_insuranceBP.setId("pane");
        Scene scene_insurance_import_S = new Scene(import_insuranceBP, 600, 400);
        scene_insurance_import_S.getStylesheets().add("ecss.css");

        //nodes for this borderpane
        TextField urlF = new TextField();
        urlF.setPrefWidth(250);
        urlF.setPrefHeight(20);

        Button search_fileB = new Button("Search Excel File");
        search_fileB.setId("button");

        Button import_excelB = new Button("Import");
        import_excelB.setId("button");

        Button back_from_importB = new Button("Back");
       back_from_importB.setId("button");

        HBox hb_nodesHB = new HBox(10);
        hb_nodesHB.setAlignment(Pos.CENTER);
        hb_nodesHB.getChildren().addAll(urlF, search_fileB, import_excelB,back_from_importB);
        import_insuranceBP.setCenter(hb_nodesHB);

        //################## action for search file button #############################
        search_fileB.setOnAction((ActionEvent t) -> {
            FileChooser fc = new FileChooser();
            FileChooser.ExtensionFilter xext = new FileChooser.ExtensionFilter("excel files (*.xls)", "*.xls");
            fc.getExtensionFilters().add(xext);
            String excel_file = fc.showOpenDialog(primaryStage).toString();
            urlF.setText(excel_file);
        });

        import_excelB.setOnAction((ActionEvent t) -> {
            if (urlF.getText().equals("")) {
                Alert warning = new Alert(AlertType.WARNING);
                warning.setTitle("Warning Dialog For Empty Text Field");
                warning.setHeaderText("File Path Is Empty");
                warning.setContentText("The File Path Is Empty.Pls Browse The Button");
                warning.showAndWait();
            } else {
                String file_path = urlF.getText();
                try {
                    Class.forName(jdbc);
                    conn = DriverManager.getConnection(url + db_name, db_user, db_pass);
                  

                    FileInputStream fin = new FileInputStream(new File(file_path));
                    XSSFWorkbook wb = new XSSFWorkbook(fin);
                    XSSFSheet sheet = wb.getSheetAt(0);
                    Row row;
                    int i;
                    for (i = 1; i <= sheet.getLastRowNum(); i++) {
                        row = sheet.getRow(i);
                        PreparedStatement ps = conn.prepareStatement("INSERT INTO clients(id,fname,lname,email)VALUES(?,?,?,?)");
                        ps.setString(1,row.getCell(0).getStringCellValue());
                        ps.setString(2, row.getCell(1).getStringCellValue());
                        ps.setString(3, row.getCell(2).getStringCellValue());
                        ps.setString(4, row.getCell(3).getStringCellValue());
                        
                        ps.execute();
                    }
                    Alert info = new Alert(AlertType.INFORMATION);
                    info.setTitle("Information Dialog.");
                    info.setHeaderText("Information Dialog For Successful Excel Data Import");
                    info.setContentText("Data From Excel Imported Successfully!");
                    info.showAndWait();
                    
                    wb.close();
                    fin.close();
                    ps.close();
                    rs.close();
                } catch (ClassNotFoundException ex) {
                    Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
                } catch (SQLException ex) {
                    Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
                }

            }

        });

        //******************** borderpane for word module***************
        BorderPane bpWord = new BorderPane();
        bpWord.setId("panew");
        Scene scenebpWord = new Scene(bpWord, 800, 400);
        scenebpWord.getStylesheets().add("ecss.css");

        HTMLEditor htmlE = new HTMLEditor();
        htmlE.setPrefHeight(250);
        htmlE.setPrefWidth(600);
        htmlE.setPadding(new Insets(75, 100, 75, 100));
        bpWord.setCenter(htmlE);

        //Buttons for scenebpWord
        Button backFromSceneWord = new Button("back");
        Button saveWordDocument = ButtonBuilder.create().text("save").build();
        //setting ids for buttons
        backFromSceneWord.setId("button");
        saveWordDocument.setId("button");
        //postioning the buttons
        HBox buttonsHB = new HBox(10);
        buttonsHB.setAlignment(Pos.BOTTOM_RIGHT);
        buttonsHB.getChildren().addAll(backFromSceneWord, saveWordDocument);
        buttonsHB.setPadding(new Insets(10, 5, 10, 5));
        bpWord.setBottom(buttonsHB);

        String written = htmlE.getHtmlText();

        Text text = TextBuilder.create().text(written).build();
        //save button for word document.
        saveWordDocument.setOnAction((ActionEvent t) -> {
            FileChooser fc = new FileChooser();
            fc.setTitle("Save Document");
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Word files (*.docx)", "*.docx");
            fc.getExtensionFilters().add(extFilter);
            File file = fc.showSaveDialog(primaryStage);

            if (file != null) {
                SaveFile(written, file);
            }
        });
        //*************** creating the button scenes **************
        BorderPane buttonsBP = new BorderPane();
        buttonsBP.setId("pane");

        Scene scene = new Scene(buttonsBP, 600, 400);
        scene.getStylesheets().add("ecss.css");

        //#########icons for buttons
        Image importIcon = new Image(getClass().getResourceAsStream("import.png"));
        ImageView importIconView = new ImageView(importIcon);
        importIconView.setFitWidth(15);
        importIconView.setFitHeight(15);

        Image exportIcon = new Image(getClass().getResourceAsStream("export.png"));
        ImageView exportIconView = new ImageView(exportIcon);
        exportIconView.setFitWidth(15);
        exportIconView.setFitHeight(15);

        Image settingsIcon = new Image(getClass().getResourceAsStream("gear-128.png"));
        ImageView settingsIconView = new ImageView(settingsIcon);
        settingsIconView.setFitWidth(15);
        settingsIconView.setFitHeight(15);

        Image wordIcon = new Image(getClass().getResourceAsStream("compose-128.png"));
        ImageView wordIconView = new ImageView(wordIcon);
        wordIconView.setFitWidth(15);
        wordIconView.setFitHeight(15);

        //creating navigation buttons
        Button oldExcelBI = new Button("Import To Table Insurance");
        Button newExcelBO = new Button("Export Insurance Records From Database");

        Button simpleWordB = new Button("Customized Word Module");
        Button settingsB = new Button("Settings");
        //setting icons for the buttons
        oldExcelBI.setGraphic(importIconView);

        newExcelBO.setGraphic(exportIconView);
        simpleWordB.setGraphic(wordIconView);
        settingsB.setGraphic(settingsIconView);

        //setting ids for buttons
        oldExcelBI.setId("button");

        newExcelBO.setId("button");
        simpleWordB.setId("button");
        settingsB.setId("button");
        //setting widths for buttons
        oldExcelBI.setPrefWidth(300);

        newExcelBO.setPrefWidth(300);
        simpleWordB.setPrefWidth(300);
        settingsB.setPrefWidth(300);
        //VBox for placing buttons
        VBox vb = new VBox(10);
        vb.getChildren().addAll(oldExcelBI, newExcelBO,settingsB);
        vb.setAlignment(Pos.CENTER);
        buttonsBP.setCenter(vb);

        //actions for buttons
        simpleWordB.setOnAction((ActionEvent t) -> {
            primaryStage.setScene(scenebpWord);
        });

        oldExcelBI.setOnAction((ActionEvent t) -> {
            primaryStage.setScene(scene_insurance_import_S);

        });
        newExcelBO.setOnAction((ActionEvent t) -> {
            primaryStage.setScene(scene_export_dbS);
        });

        //program buttons
        backFromSceneWord.setOnAction((ActionEvent t) -> {
            primaryStage.setScene(scene);
        });
        back_from_exportS.setOnAction((ActionEvent t) -> {
             primaryStage.setScene(scene);
        });
        back_from_importB.setOnAction((ActionEvent t) -> {
             primaryStage.setScene(scene);
        });
        //primaryStage
        Image stageIcon = new Image(getClass().getResourceAsStream("tools-128.png"));
        primaryStage.getIcons().add(stageIcon);
        primaryStage.setScene(scene);
        primaryStage.setResizable(false);
        primaryStage.setTitle("Excel Productivity Tool");
        primaryStage.show();
    }

    private void SaveFile(String content, File file) {

        try {
            FileWriter fileWriter = null;
            fileWriter = new FileWriter(file);
            fileWriter.write(content);
            fileWriter.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelProductivity.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

}
