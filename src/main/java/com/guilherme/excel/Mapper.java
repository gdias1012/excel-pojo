package com.guilherme.excel;

import com.guilherme.excel.annotation.Celula;
import com.guilherme.excel.annotation.Planilha;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.lang.reflect.Field;
import java.util.*;

public class Mapper {

    private Map<String, Integer> columnsMap;

    public Object readObject(String path, Class<?> objectType) throws Exception {
        String sheetName = objectType.getAnnotation(Planilha.class).nome();
        Sheet s = openSheet(path, sheetName);
        loadColumnNames(s);
        return readObject(s.getRow(0), objectType);
    }

    public Object readObject(Row r, Class<?> objectType) throws Exception {
        Object o = objectType.newInstance();
        Field[] fields = objectType.getDeclaredFields();
        for(Field f : fields) {
            f.setAccessible(true);
            String cellName = f.getAnnotation(Celula.class).nome();
            f.set(o, getColumnValue(r, cellName));
            f.setAccessible(false);
        }
        return o;
    }

    public List readCollection(String path, Class<?> type) throws Exception {
        List<Object> l = new ArrayList();
        String sheetName = type.getAnnotation(Planilha.class).nome();
        Sheet s = openSheet(path, sheetName);
        loadColumnNames(s);
        Iterator<Row> iterator = s.rowIterator();
        iterator.next();
        while(iterator.hasNext()) {
            Row r = iterator.next();
            l.add(readObject(r, type));
        }
        return l;
    }

    private Sheet openSheet(String path, String sheetName) throws Exception {
        File f = new File(path);
        Workbook w = WorkbookFactory.create(f);
        return w.getSheet(sheetName);
    }

    private void loadColumnNames(Sheet s) {
        Row r = s.getRow(0);
        columnsMap = new HashMap<>();
        r.cellIterator().forEachRemaining(x -> columnsMap.put(x.getStringCellValue(), x.getColumnIndex()));
    }

    private String getColumnValue(Row r, String name) {
        Cell c = r.getCell(columnsMap.get(name));
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(c);
    }

}
