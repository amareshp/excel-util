package com.visa.test;

import com.visa.test.util.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.util.*;

/**
 * Unit test for simple App.
 */

@org.testng.annotations.Test
public class AppTest {
    private final Logger logger = LoggerFactory.getLogger(ExcelReader.class);
    Map<String, String> childParentMap = new HashMap<String, String>();
    Map<String, Set<String>> bidMap = new LinkedHashMap<String, Set<String>>();

    @Test
    public void excelReadTest() throws Exception{
        ExcelReader excelReader = new ExcelReader("src/test/resources/Playbook_Accounts_Account_Hierarchy_Bids_041416.xlsx");
        List<Row> rows = excelReader.getAllExcelRows("Sales Hierarchy File", false);

        //create the child parent map
        for(Row row : rows) {
            Cell bidCell = row.getCell(1);
            Cell parentBidCell = row.getCell(4);
            if(bidCell != null && parentBidCell != null) {
                bidCell.setCellType(Cell.CELL_TYPE_STRING);
                parentBidCell.setCellType(Cell.CELL_TYPE_STRING);
                childParentMap.put(bidCell.getStringCellValue(), parentBidCell.getStringCellValue());
            }
        }

        for(Row row : rows) {
            Cell acctName = row.getCell(0);
            Cell bidCell = row.getCell(1);
            Cell countryCell = row.getCell(2);
            Cell parentAcct = row.getCell(3);
            Cell parentBidCell = row.getCell(4);
            if(bidCell != null) {
                bidCell.setCellType(Cell.CELL_TYPE_STRING);
                String bidValue = bidCell.getStringCellValue();
                Set<String> parentSet = bidMap.get(bidValue);

                if(parentSet == null) {
                    Set<String> newParentSet = new HashSet<String>();
                    if(parentBidCell != null) {
                        parentBidCell.setCellType(Cell.CELL_TYPE_STRING);
                        String parentBidValue = parentBidCell.getStringCellValue();
                        newParentSet.add(parentBidValue);
                        Set<String> moreParents = getParentSet(parentBidValue, new HashSet<String>());
                        if(moreParents != null) {
                            newParentSet.addAll(moreParents);
                        }
                    }
                    bidMap.put(bidValue, newParentSet);
                } else {
                    if(parentBidCell != null) {
                        parentBidCell.setCellType(Cell.CELL_TYPE_STRING);
                        String parentBidValue = parentBidCell.getStringCellValue();
                        parentSet.add(parentBidValue);
                        Set<String> moreParents = getParentSet(parentBidValue, new HashSet<String>());
                        if(moreParents != null) {
                            parentSet.addAll(moreParents);
                        }
                        bidMap.put(bidValue, parentSet);
                    }
                }

            }
        }
        printMap(bidMap);
    }

    private Set<String> getParentSet(String child, Set<String> result) {
        if(StringUtils.isEmpty(child)) { return result; }
        if(StringUtils.isEmpty(child.trim())) { return result; }

        String parent = childParentMap.get(child);
        if( StringUtils.isNotEmpty(parent) && !parent.trim().toLowerCase().equals(child.trim().toLowerCase()) ) {
            result.add(parent);
            getParentSet(parent, result);
        }

        return result;
    }

    private void printMap(Map<String, Set<String>> map) {
        for(String key : map.keySet()) {
            logger.info(key + " : " + StringUtils.join(map.get(key).iterator(), ",") );
        }
    }

}
