package za.co.silversurge.dataprocessing;

import org.apache.poi.ss.usermodel.Row;

public interface ExcelOnRowCreateListener {

    void onRowCreate(int index, int max, Row row);

}
