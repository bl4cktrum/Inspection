import java.io.*;
import javax.swing.*;

/**
 *
 * @author BL4CKTRUM
 */
public class InspectionGUI {
    public static MainFrame mf;
    public static String filepath;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        showScreen();
    }

    private static void showScreen(){

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); //set look&feel to system props.

                    JFileChooser chooser = new JFileChooser();
                    chooser.showOpenDialog(null);
                    File f = chooser.getSelectedFile();
                    filepath = f.getAbsolutePath();
                    System.out.println(filepath);

                    mf = new MainFrame();
                    mf.setVisible(true);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }
}

/*
public class Test {

    public static void main(String[] args) throws InvalidFormatException, IOException {
        File file = new File("C:/Users/yavuz/IdeaProjects/inspection/src/main/java/inspection.xlsx");
        OPCPackage pkg = OPCPackage.open(new FileInputStream(file));
        XSSFWorkbook wb = new XSSFWorkbook(pkg);
        //
        int finding = 445;

        boolean write = false;

        DataFormatter formatter = new DataFormatter();
        for(Sheet sheet : wb) {
            for(Row row : sheet){
                if(row.getCell(0)!=null && !formatter.formatCellValue(row.getCell(0)).equals("")){
                    Cell cell = row.getCell(0);
                    String text = formatter.formatCellValue(cell);
                    if('0'<=text.charAt(0) && text.charAt(0)<='9') {
                        int id = Integer.parseInt(text);
                        if (id == finding) {
                            System.out.println(sheet.getSheetName());
                            System.out.println(sheet.getRow(row.getRowNum()).getCell(1));
                            Cell cellCurrent = row.getCell(2);
                            if (cellCurrent == null){
                                cellCurrent = row.createCell(2);
                            }
                            cellCurrent.setCellValue("X");
                            write = true;
                        }
                    }
                }
            }
        }

        if (write) {
            System.out.println("writing");
            FileOutputStream outputStream = new FileOutputStream(file);
            wb.write(outputStream);
            outputStream.close();
            wb.close();
        }
    }
}

 */
