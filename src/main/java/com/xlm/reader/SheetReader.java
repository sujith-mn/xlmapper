package com.xlm.reader;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xlm.util.AnnotationUtil;

/**
 * @author sujithna
 * @param <T>
 *
 */
public class SheetReader<T> {

  private final DataFormatter df;
  private final FormulaEvaluator evaluator;

  private int headerRow = 0;
  public void setHeaderRow(int headerRow) {
    this.headerRow = headerRow;
  }

  public SheetReader(Workbook workbook) {
    this.df = new DataFormatter();
    evaluator = workbook.getCreationHelper().createFormulaEvaluator();
  }

  public List<T> getListFromSheet(Sheet datatypeSheet, Class<T> claz) throws Exception {

    List<T> rows = new ArrayList<T>();
    Map<Integer, Method> headerPositionWithSetterMethods = getHeaderPositionAndSetterMethods(
            datatypeSheet.getRow(this.headerRow).iterator(),
            AnnotationUtil.getFields(claz)
    );

    if(headerPositionWithSetterMethods.size()==0){
      return rows;
    }

    for (Row currentRow : datatypeSheet) {
      if (currentRow.getRowNum() <= this.headerRow) {
        continue;
      }
      T pojoInstance = bind(currentRow, claz,headerPositionWithSetterMethods);
      rows.add(pojoInstance);
    }
    return rows;
  }

  private T bind(Row row, Class<T> claz,Map<Integer, Method> headerPositionWithSetterMethods )
          throws Exception {
    T pojoInstance = claz.getDeclaredConstructor().newInstance();
    for (int position : headerPositionWithSetterMethods.keySet()) {
      Cell currentCell = row.getCell(position);
      if (currentCell!= null) {
        setCellValueToField(pojoInstance, headerPositionWithSetterMethods.get(position), currentCell);
      }
    }
    return pojoInstance;
  }

  private void setCellValueToField(T pojoInstance, Method fieldSetterMethod, Cell currentCell)
          throws Exception {
    if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
      fieldSetterMethod.invoke(pojoInstance, this.evaluator.evaluate(currentCell).getNumberValue() + "");
    }
    fieldSetterMethod.invoke(pojoInstance, this.df.formatCellValue(currentCell));
  }

  private Map<Integer, Method> getHeaderPositionAndSetterMethods(Iterator<Cell> cellIterator,
                                                                 Map<String, Method> fieldSetterMethods) {
    Map<Integer, Method> headerPositionWithSetterMethods = new HashMap<Integer, Method>();
    int cellPosition = 0;
    while (cellIterator.hasNext()) {
      String fieldName = this.df.formatCellValue(cellIterator.next()).trim()
              .replaceAll("[\\t\\n\\r]+", " ");
      Method fieldSetterMethod = fieldSetterMethods.get(fieldName);
      if (fieldSetterMethod != null) {
        headerPositionWithSetterMethods.put(cellPosition, fieldSetterMethod);
      }
      cellPosition++;
    }
    return headerPositionWithSetterMethods;
  }

}
