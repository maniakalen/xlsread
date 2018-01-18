package maniakalen.xls;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.ArrayUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.*;

public class Reader {


    private static final int NUMBER_INT = 0;
    private static final int NUMBER_DOUBLE = 1;
    //private static final String FILE_NAME = "C:/Users/peter.georgiev/Documents/test.xlsx";
    //private static  String[] sheetsList = {"B.case", "B.Case PS"};

    private Connection conn;

    private String[] sheetsList;

    public static void main(String[] args) {
        (new Reader()).run();
    }
    public void run()
    {
        if (prepareDbConnection()) {
            try {
                List<Map<String, Object>> records = getPendingRecords();
                for (Map<String, Object> record : records) {
                    File file = new File(record.get("file_path").toString());
                    if (!file.exists()) {
                        throw new FileNotFoundException("File not found: " + record.get("file_path").toString());
                    }

                    workOnRecord(
                            record.get("offer_id").toString(),
                            record.get("object_type").toString(),
                            Integer.parseInt(record.get("object_id").toString()),
                            record.get("sheets_list").toString().split("\\|"),
                            file,
                            Integer.parseInt(record.get("record_id").toString())
                    );
                }
                conn.close();
            } catch (Exception ex) {
                ex.printStackTrace(System.err);
            }
        }
    }
    private Properties getConfigProperties() throws IOException
    {
        Properties prop = new Properties();
        InputStream input = null;

        try {

            input = new FileInputStream("config.properties");

            prop.load(input);
            return prop;

        } catch (IOException ex) {
            ex.printStackTrace(System.err);
            throw ex;
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    private boolean prepareDbConnection()
    {
        Properties config = null;
        try {
            config = getConfigProperties();
            Class.forName("com.mysql.jdbc.Driver").newInstance();
        } catch (Exception ex) {
            ex.printStackTrace(System.err);
            return false;
        }
        String host = System.getProperty("db.host", config.getProperty("db.host", ""));
        String schema = System.getProperty("db.schema", config.getProperty("db.schema", ""));
        String user = System.getProperty("db.user", config.getProperty("db.user", ""));
        String pass = System.getProperty("db.pass", config.getProperty("db.pass", ""));
        if (host.equals("") || schema.equals("") || user.equals("") || pass.equals("")) {
            return false;
        }
        try {
            conn = DriverManager.getConnection("jdbc:mysql://" + host + "/" + schema + "?" +
                            "user=" + user + "&password=" + pass);
            return true;
        } catch (SQLException ex) {

            ex.printStackTrace(System.err);
            return false;
        }
    }

    private boolean printSheetData(String offerId, String objectType, int objectId, Sheet dataSheet)
    {
        Iterator<Row> iterator = dataSheet.iterator();
        //HashMap<String, Object> cellData = new HashMap<String, Object>();
        String insertSt = "INSERT INTO excel_import(id_propuesta_comercial, object_type, " +
        "object_id, import_date, sheet_index, sheet_name, column_code, row_index, value, valueComputed)" +
                "VALUES('" + offerId + "','" + objectType + "'," + objectId + ",NOW(),?,?,?,?,?,?)";
        PreparedStatement stmt = null;
        try {
            boolean lastAutoCommitStatus = conn.getAutoCommit();
            conn.setAutoCommit(false);
            int sheetIndex = dataSheet.getWorkbook().getSheetIndex(dataSheet);
            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();

                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    stmt = conn.prepareStatement(insertSt);
                    stmt.setInt(1, sheetIndex);
                    stmt.setString(2, dataSheet.getSheetName());
                    stmt.setString(3, ((char)(currentCell.getColumnIndex() + 65)) + "");
                    stmt.setInt(4, (int) (currentCell.getRowIndex() + 1));
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        stmt.setString(5, currentCell.getStringCellValue());
                        stmt.setString(6, currentCell.getStringCellValue());
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        setStatementNumericParam(stmt, 5, currentCell.getNumericCellValue());
                        setStatementNumericParam(stmt, 6, currentCell.getNumericCellValue());
                        //stmt.setDouble(5, currentCell.getNumericCellValue());
                        //stmt.setDouble(6, currentCell.getNumericCellValue());
                    } else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
                        stmt.setString(5, currentCell.getCellFormula());
                        switch (currentCell.getCachedFormulaResultType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                setStatementNumericParam(stmt, 6, currentCell.getNumericCellValue());
                                break;
                            case Cell.CELL_TYPE_STRING:
                                stmt.setString(6, currentCell.getStringCellValue() + "");
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                stmt.close();
                                continue;
                            case Cell.CELL_TYPE_BOOLEAN:
                                stmt.setString(6, Boolean.toString(currentCell.getBooleanCellValue()));
                            case Cell.CELL_TYPE_ERROR:
                                stmt.setString(6, "#ERROR");

                        }

                    } else if (currentCell.getCellTypeEnum() == CellType.BLANK) {
                        //stmt.setString(5, "");
                        //stmt.setString(6, "");
                        stmt.close();
                        continue;
                    } else if (currentCell.getCellTypeEnum() == CellType.BOOLEAN) {
                        stmt.setString(5, Boolean.toString(currentCell.getBooleanCellValue()));
                        stmt.setString(6, Boolean.toString(currentCell.getBooleanCellValue()));
                    }
                    stmt.execute();
                    stmt.close();
                }
            }
            conn.commit();
            conn.setAutoCommit(lastAutoCommitStatus);
        } catch (SQLException ex) {
            System.out.println(stmt);
            ex.printStackTrace(System.err);
            return false;
        }
        return true;
    }
    private void setStatementNumericParam(PreparedStatement stmt, int param, Number value)
    {
        try {
            if (getNumberType(value) == NUMBER_INT) {
                stmt.setInt(param, value.intValue());
            } else {
                stmt.setDouble(param, value.doubleValue());
            }
        } catch (SQLException ex) {
            ex.printStackTrace(System.err);
        }
    }
    private int getNumberType(Number number)
    {
        int hole = number.intValue();
        if ((number.doubleValue() % hole) > 0) {
            return NUMBER_DOUBLE;
        }
        return NUMBER_INT;
    }
    private void workOnRecord(String offerId, String objectType, int objectId, String[] sheetsList, File file, int rec_id)
    {
        try {
            FileInputStream excelFile = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(excelFile);
            boolean done = sheetsList.length > 0;
            for (String sheet : sheetsList) {
                try {
                    done = printSheetData(offerId, objectType, objectId, workbook.getSheet(sheet)) && done;
                } catch (NullPointerException ex) {
                    ex.printStackTrace(System.err);
                }
            }
            if (done) {
                try {
                    PreparedStatement stmt = conn.prepareStatement("UPDATE excel_import_pending_files SET status_work = 'done' WHERE id = ?");
                    stmt.setInt(1, rec_id);
                    stmt.executeUpdate();
                    stmt.close();
                } catch (SQLException ex) {
                    System.err.println("Unable to update status");
                    ex.printStackTrace(System.err);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace(System.err);
        } catch (IOException e) {
            e.printStackTrace(System.err);
        }
    }
    private List<Map<String, Object>> getPendingRecords() throws Exception
    {
        Statement stmt = null;
        try {
            stmt = conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,
                    ResultSet.CONCUR_UPDATABLE);
            ResultSet uprs = stmt.executeQuery(
                    "SELECT * FROM excel_import_pending_files where status_work = 'pending'");
            List<Map<String, Object>> results = new ArrayList<Map<String, Object>>();
            Map<String, Object> record = null;
            while (uprs.next()) {
                record = new HashMap<String, Object>();

                record.put("record_id", uprs.getInt("id"));
                record.put("offer_id", uprs.getString("id_propuesta_comercial"));
                record.put("object_type", uprs.getString("object_type"));
                record.put("object_id", uprs.getInt("object_id"));
                record.put("sheets_list", uprs.getString("sheets_list"));
                record.put("file_path", uprs.getString("file_path"));

                results.add(record);

                uprs.updateString( "status_work", "in_progress");
                uprs.updateRow();
            }

            uprs.close();
            stmt.close();
            return results;
        } catch (Exception e ) {
            e.printStackTrace(System.err);
            throw e;
        }
    }
}
