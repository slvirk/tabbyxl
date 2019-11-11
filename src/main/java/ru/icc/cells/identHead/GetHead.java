package ru.icc.cells.identHead;


import org.apache.commons.lang.ObjectUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import ru.icc.cells.ssdc.model.CCell;
import ru.icc.cells.ssdc.model.CTable;
import ru.icc.cells.ssdc.model.style.BorderType;
import ru.icc.cells.ssdc.model.style.CBorder;
import ru.icc.cells.ssdc.model.style.CStyle;

import java.awt.geom.NoninvertibleTransformException;
import java.io.Console;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Stack;


public class GetHead {
    CTable table;
    int hB =0, hR = 0;
    CellPoint cellShift;
    WorkbookManage workbookManage;
    enum CellParam { WIDTH, HEIGHT, BOTH}

    private int tmpC =0;


    public GetHead(CTable inputTable, int [] shift, Workbook workbook, String sheetName){
        int c=0, cId;
        table = inputTable;
        cellShift = new CellPoint(shift);
        hR = table.numOfCols();
        workbookManage = new WorkbookManage(workbook, sheetName);
        System.out.println("GetHead constructor in the progress");
        System.out.println(String.format("%s sheet is processing", sheetName));
        //First cell determine the border of the header
        CCell cell = getCellByCoord(1, 1);

        if (cell.getRb() != cell.getRt())
            hB = cell.getRb();
        else
            hB = getHeaderLine();

        System.out.printf("Head bottom = %d\n", hB);
        System.out.printf("Head right = %d\n", hR);
        System.out.printf("Value '%s'\n", getCellByCoord(1,1).getText());

    }

    private boolean isBorder(CCell cell, boolean isBottom){
        /*cell - cell for analyzis
        posititon: true - bottom line; false - top line.
        */
        BorderType borderType;
        if (isBottom == true)
            borderType = cell.getStyle().getBottomBorder().getType();
        else
            borderType = cell.getStyle().getTopBorder().getType();
        if (borderType == BorderType.NONE)
            return false;
        else
            return true;
    }

    private boolean checkBorderRow(CCell cell, boolean isBottom){
        CCell nextCell;
        boolean ib=true;

        do{
            if ( cell == null ){
                ib = false;
                break;
            }

            if (isBorder(cell, isBottom)){
                cell = getCellByCoord(cell.getCr()+1, cell.getRt());
            }
            else
                ib = false;

        }while ((ib) && (cell.getCr()<=hR));
        return ib;

    }

    private int getHeaderLine(){
        CCell curCell, nextCell;
        int res = 0;
        curCell = getCellByCoord(1, 1);
        do{
            if ( curCell == null ) return  0;
            if (isBorder(curCell,true)){
                res = curCell.getRb();
            }
            else{
                curCell = getCellByCoord(curCell.getCl(), curCell.getRb() +1);

            }
        }while (res == 0);
        return res;
    }



    public void analyzeHead(){
        CellCoordNode cellCoordNode;
        CCell cCell, tmpCell;
        Block headBlock;
        int b,r;
        boolean lbl;
        int curCellTop = 1, curCellLeft = 1;
        do{
            //Get top level block borders
            cCell = getCellByCoord(curCellLeft, 1);
            if (! isLabel(cCell)) {
                //Get lower cell size
                cCell = expCell(cCell, hR, hB);

                headBlock = new Block(cCell);
                if ((! isLabel(cCell)) && (cCell.getStyle().getLeftBorder().getType() != BorderType.NONE) && (cCell.getCr() < hR)) {

                    tmpCell = getCellByCoord(cCell.getCr()+1, cCell.getRt());
                    if ((tmpCell != null)) {
                        cCell = expByHeight(cCell, tmpCell.getRb());
                        if (cCell.getRb() == tmpCell.getRb()) {
                            headBlock.setRight(tmpCell.getCr());
                            headBlock.setBottom(cCell.getRb());
                        }
                    }
                }

                cCell = cellTransofrm(cCell, headBlock);

            }
            else{
                tmpCell = checkRight(cCell);
                tmpCell = expByHeight(tmpCell); //cCell or tmpCell
                if (tmpCell.equals(cCell))
                    cCell = tmpCell;
                else
                    cCell = cellTransofrm(tmpCell, new Block(tmpCell));

            }

            System.out.println(String.format("Cell block(l:%s, r:%s, t:%s, b:%s) Value=%s", cCell.getCl(), cCell.getCr() ,cCell.getRt(), cCell.getRb(), cCell.getText()));

            if (cCell.getRb() < hB)
                buildBlock(cCell);
            //Value for next cell
            curCellLeft = cCell.getCr() + 1;
        }while(curCellLeft <= hR);
    }


    private CCell expCell(CCell cCell, int rightBorder, int bottomBorder){
        Block block;
        CCell tmpCell;
        do {
            tmpCell = cCell;
            block = new Block(cCell);
            cCell = expByHeight(cCell, bottomBorder);
            if (isLabel(cCell))
                return cCell;
            cCell = expByWidth(cCell, rightBorder);
            if (block.equals(new Block(cCell)))
                break;
            if (tmpCell.equals(cCell))
                break;
        }while (cCell.getText().isEmpty());

        return cCell;
    }

    private CCell expByWidth(CCell emptyCell){
        return expByWidth(emptyCell, emptyCell.getCr()+1);
    }
    private CCell expByWidth(CCell emptyCell, int rightBorder){
        CCell nextCell, tmpCell = emptyCell;
        boolean next = true;

        if (emptyCell.getStyle().getRightBorder().getType() != BorderType.NONE)
            next = false;

        while (next){
            if ((emptyCell.getCr()+1)> rightBorder)
                break;
            nextCell = getCellByCoord(emptyCell.getCr()+1, emptyCell.getRt());
            if (nextCell == null)
                return emptyCell;
            //Check
            //nextCell = expByHeight(nextCell, emptyCell.getRb());
            if ((nextCell == null) || (nextCell.getStyle().getLeftBorder().getType() != BorderType.NONE))
                return emptyCell;
            if (nextCell.getCr() >= rightBorder)
                next = false;
            if ( emptyCell.getRb() == nextCell.getRb() ){
                if (! isLabel(emptyCell)){
                    //It is possible to merge
                    nextCell.merge(emptyCell);
                    emptyCell = nextCell;

                    if ((isLabel(emptyCell)) || (emptyCell.getStyle().getRightBorder().getType() != BorderType.NONE)) //Cell has text
                        next = false;
                }
                else
                    next = false;
            }
            else next=false;

        };
        /*
        if (! isLabel(emptyCell)){
            emptyCell= mergeToLeft(emptyCell);
            System.out.println(String.format("Empty cell is (l:%s, r:%s, t:%s, b:%s)", emptyCell.getCl(), emptyCell.getCr(), emptyCell.getRt(), emptyCell.getRb()));
        }
        */
        if (!emptyCell.equals(tmpCell))
            cellTransofrm(emptyCell, new Block(emptyCell));
        return emptyCell;

    }

    private CCell expByWidth_(CCell emptyCell, int rightBorder){
        CCell nextCell, bottomCell;

        bottomCell = getCellByCoord(emptyCell.getCl(), emptyCell.getRb()+1);

        //Check the border inside
        if (emptyCell.getStyle().getRightBorder().getType() != BorderType.NONE)
            return emptyCell;
        while (emptyCell.getCr() < rightBorder){
            nextCell = getCellByCoord(emptyCell.getCr()+1, emptyCell.getRt());
            if (nextCell.getRb() < emptyCell.getRb())
                    nextCell = expByHeight(nextCell, emptyCell.getRb());
            //Check border outside
            if ((nextCell == null) || (nextCell.getStyle().getLeftBorder().getType() != BorderType.NONE))
                return emptyCell;
            //Check height of the cells
            if (nextCell.getRb() != emptyCell.getRb()) return emptyCell;
            //if any border
            if ((emptyCell.getStyle().getRightBorder().getType() == BorderType.NONE) &&
                (nextCell.getStyle().getLeftBorder().getType() == BorderType.NONE)){
                //Merging cell
                nextCell.merge(emptyCell); //Try it
                //mergeCells(nextCell, emptyCell);
                if (isLabel(nextCell)){
                    return nextCell;
                }

                emptyCell = nextCell;

            }
            else
                break;
        }
        if ((bottomCell != null) && (isLabel(bottomCell) && ! isLabel(emptyCell)) && (bottomCell.getCr() == emptyCell.getCr()))
            emptyCell = expByHeight(emptyCell, bottomCell.getRb());
        //emptyCell = checkRight(emptyCell);
        return emptyCell;
    }

    private CCell expByHeight(CCell emptyCell){
        return expByHeight(emptyCell, -1);
    }

    private CCell expByHeight(CCell emptyCell, int bottomBorder){
        boolean isMerge;
        CCell nextCell, tmpCell = emptyCell;
        String cellVal ="";

        //if ((emptyCell.getCl() == 3) && (emptyCell.getCr()== 7))
        //    System.out.println("1111");
        //Set bottom as a bottom border of header
        if (bottomBorder == -1 )
            bottomBorder = hB;

        if (emptyCell.getRb() == hB)
            return emptyCell;

        do{
            isMerge = true;
            if (bottomBorder > hB)
                bottomBorder = hB;
            //Check lower border
            if ((emptyCell.getRb() == bottomBorder) || (emptyCell.getStyle().getBottomBorder().getType() != BorderType.NONE))
                return  emptyCell;

            //Get lower cell
            nextCell = getCellByCoord(emptyCell.getCl(), emptyCell.getRb() + 1);
            if ( nextCell == null)
                break;
            //nextCell = expByHeight(nextCell, bottomBorder);
            nextCell = expByWidth(nextCell, emptyCell.getCr());
            if (emptyCell.getCr() != nextCell.getCr()) {

                return emptyCell; //if width of lower cell doesn't equial that current
            }
            if (emptyCell.getStyle().getBottomBorder().getType() != BorderType.NONE)
                isMerge = false;
            if (nextCell.getStyle().getTopBorder().getType() != BorderType.NONE)
                isMerge = false;

            if (isMerge){
                        //Merging cell
                        cellVal = emptyCell.getText();
                        nextCell.merge(emptyCell);
                        if (!nextCell.getText().isEmpty())
                            cellVal = (! cellVal.isEmpty()) ?  cellVal + "/n" + nextCell.getText() : nextCell.getText();
                        nextCell.setText(cellVal);
                        emptyCell = nextCell;

            }
            else
                break;

        }while (emptyCell.getRb() < bottomBorder);
        emptyCell= expByWidth(emptyCell, emptyCell.getCr());
        if (!emptyCell.equals(tmpCell))
            cellTransofrm(emptyCell, new Block(emptyCell));
        //emptyCell = checkLower(emptyCell, bottomBorder);
        return emptyCell;
    }

    void buildBlock(CCell topCell){
        buildBlock(topCell, null);
    }

    void buildBlock(CCell topCell, Block block){
        CCell tmpCell;
        Block tmpBlock;
        boolean direction = true; //True - go downwards; false - go upwards
        Stack<CCell> blockItems = new Stack<>();
        //if (! isLabel(topCell)) topCell = expByWidth(topCell);

        CCell curCell, newCell = null;
        if (block == null) {
            block = new Block(topCell); //Default block size
            block.setBottom(hB);
        }
        if (! isLabel(topCell)) topCell = expCell(topCell, block.getRight(), block.getBottom());
        blockItems.push(topCell);
        while(! blockItems.empty()){
            curCell = blockItems.peek();
            if( (curCell.getRb() == hB) && (direction)) direction = false;

            if (direction == true){
                newCell = getCellByCoord(curCell.getCl(), curCell.getRb()+1);
                if ((newCell != null) && (! isLabel(newCell))){
                    System.out.print(String.format("! cell_old(l:%s, r:%s, t:%s, b:%s) - ", newCell.getCl(), newCell.getCr(), newCell.getRt(), newCell.getRb()));
                    Block newBlock = new Block(newCell);
                    newBlock.setBottom(block.getBottom());
                    newBlock.setRight(curCell.getCr());
                    newCell = cellTransofrm(newCell, newBlock);
                    tmpCell = getUpperCell(newCell.getCl(), newCell.getRt());

                    if ((tmpCell != null) && (isLabel(tmpCell)) && (tmpCell.getCl() == newCell.getCl()) &&
                            (tmpCell.getCr() == newCell.getCr()) && (tmpCell.getRb() + 1 == newCell.getRt())
                            && (! isLabel(tmpCell) && (tmpCell.getStyle().getBottomBorder().getType() == BorderType.NONE))
                            && (newCell.getStyle().getTopBorder().getType() == BorderType.NONE)
                    ){
                        tmpBlock = new Block(tmpCell);
                        tmpBlock.setBottom(newCell.getRb());
                        newCell = cellTransofrm(tmpCell, tmpBlock);
                        blockItems.pop();
                    }
                    System.out.println(String.format("cell_new(l:%s, r:%s, t:%s, b:%s) Value = '%s'", newCell.getCl(), newCell.getCr(), newCell.getRt(), newCell.getRb(), newCell.getText()));
                }

                else if(newCell == null) {
                    newCell = expByHeight(curCell);//mergeVertCells(newCell);
                }
                else if ( (newCell != null) ){
                    newCell = checkRight(newCell, block.getRight());
                }


                blockItems.push(newCell);
                block.incSizeH(newCell);
                if (newCell.getRb() == hB) direction = false;
            }
            else{
                //Back step
                newCell = blockItems.pop();
                //if there are no any items in stack than exit
                if (blockItems.empty()) break;
                curCell = blockItems.peek();

                if (newCell.getCr() < curCell.getCr()){
                    //sub column
                    newCell = getCellByCoord(newCell.getCr()+1, newCell.getRt());
                    if (newCell == null) break;
                    if (! isLabel(newCell)){
                        System.out.print(String.format("- cell_old(%s, %s, %s, %s) - ", newCell.getCl(), newCell.getCr(), newCell.getRt(), newCell.getRb()));
                        newCell = cellTransofrm(newCell, block);
                        System.out.println(String.format("cell_new(%s, %s, %s, %s), Value = %s", newCell.getCl(), newCell.getCr(), newCell.getRt(), newCell.getRb(), newCell.getText()));
                    }
                    if ((newCell.getRb() == hB) && (newCell.getCr() == hR)) break;
                    //Check possibility of transformation
                    //curCell = expByHeight(newCell);
                    curCell = expCell(newCell, block.getRight(), block.getBottom());

                    if (curCell.getRb() != newCell.getRb())
                        newCell = cellTransofrm(curCell, new Block(curCell));
                    else
                        newCell = curCell;
                    blockItems.push(newCell);
                    direction = true;
                }
            }

        }

    }

    private boolean isLabel(CCell cell){
        if ( cell == null ) return false;
        if (cell.getText() == "")
            return false;
        else
            return true;
    }

    private CCell cellTransofrm(CCell cCell,  Block block){
        CCell neighborCellR = null, neighborCellB = null;
        boolean f;
        if ((cCell.getCl()==4) && (cCell.getCr() == 4))
            System.out.println("!!!!!!");
        block.incSizeH(cCell);
        do{
            f = false;
            if (cCell.getRb() < block.getBottom() ){
                neighborCellB = expByHeight(cCell, block.getBottom());
                if ((neighborCellB != null) && (cCell.getRb() < neighborCellB.getRb())){
                    f = true;
                    cCell =neighborCellB;
                }
            }

            if (cCell.getCr() < block.getRight()){
                //Try to expand to width
                neighborCellR = expByWidth(cCell, block.getRight());
                //neighborCellR = checkRight(neighborCellR);
                if ((neighborCellR != null) && (cCell.getCr()< neighborCellR.getCr())){
                    f = true;
                    cCell = neighborCellB;

                }
            }

        }while (f);

        //Expansion in width
        workbookManage.mergeCells(new Block(cCell), cellShift, tmpC++);

        return  cCell;
    }

    private CCell _cellTransofrm(CCell cCell,  Block block){
        //Old function of transformation
        //should be deleted
        int expBottom;
        Block blockForMerge;
        CCell neighborCellR = null, neighborCellB = null;

        cCell = expByWidth(cCell);

        //Get the expected width
        if (cCell.getCr() + 1 <= block.getRight()) {
            neighborCellR = getCellByCoord(cCell.getCr() + 1, cCell.getRt());
        }
        expBottom = (neighborCellR == null) ? cCell.getRb()  : neighborCellR.getRb();
        cCell = expByHeight(cCell, expBottom );

        if ((neighborCellR != null) && (neighborCellR.getRt()<=block.getRight()) && (equals(cCell, neighborCellR, CellParam.HEIGHT)))
            cCell= neighborCellR.merge(cCell);

        if ((neighborCellB != null) && (isLabel(neighborCellB)) && (equals(cCell, neighborCellB, CellParam.WIDTH)))
            cCell= neighborCellR.merge(cCell);
            //cCell = mergeCells(neighborCellR, cCell);
        blockForMerge = new Block(cCell);
        workbookManage.mergeCells(blockForMerge, cellShift, tmpC++);

        return cCell;
    }

    private boolean equal(CCell cell1, CCell cell2){
        return equals(cell1, cell2, CellParam.BOTH);
    }

    private boolean equals(CCell cell1, CCell cell2, CellParam cellParam){
        boolean resultW = false, resultH = false;
        if ((cell1 == null) || (cell2 == null))
            return false;
        if ((cell1.getCl() == cell2.getCl()) && (cell1.getCr() == cell2.getCr())) resultW = true;
        if (cellParam == CellParam.WIDTH) return resultW;
        if ((cell1.getRt() == cell2.getRt()) && (cell1.getRb() == cell2.getRb())) resultH = true;
        if (cellParam == CellParam.HEIGHT) return resultH;
        return resultH & resultW;
    }

    public CCell getCellByCoord(int l, int t){
        //if (t > hB ) t = hB;
        Iterator<CCell> cCellIterator = table.getCells();
        while (cCellIterator.hasNext()){
            CCell cCell = cCellIterator.next();
            if ((cCell.getCl() == l) && (cCell.getRt() == t)){
                return cCell;
            }
        }
        return null;
    }

    public CCell getUpperCell(int l, int b){
        Iterator<CCell> cCellIterator = table.getCells();
        while (cCellIterator.hasNext()){
            CCell cCell = cCellIterator.next();
            if ((cCell.getCl() == l) && (cCell.getRb() + 1 == b)){
                if (cCell.getStyle().getBottomBorder().getType() != BorderType.NONE)
                    return null;
                else
                    return cCell;
            }
        }
        return null;
    }

    public CCell getLeftCell(CCell curCell){
        int r, t;
        r = curCell.getCl() - 1;
        t = curCell.getRt();
        Iterator<CCell> cCellIterator = table.getCells();
        while (cCellIterator.hasNext()) {
            CCell cCell = cCellIterator.next();
            if (( cCell.getCr() == r ) && ( cCell.getRt() == t )){
                return cCell;
            }
        }
        return null;
    }

    public void saveWorksheet(String path){
        try {
            workbookManage.saveWorkbook(path);
            System.out.println("Copy was save");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void saveWorkbook(String path){
        try {
            workbookManage.saveWorkbook(path);
            System.out.println("Copy was save");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private CCell checkLower(CCell curCell){ return checkLower(curCell, hB);}
    private CCell checkLower(CCell curCell, int bottomBorder){
        CCell resCell = curCell,
              tmpCell = null;
        boolean next = true;
        int ul = resCell.getRt();;
        do{
           if (resCell.getRb() +1 <= bottomBorder){
               if (resCell.getStyle().getBottomBorder().getType() != BorderType.NONE) {
                   next = false;
                   break;
               }
               tmpCell= getCellByCoord(resCell.getCr(), resCell.getRb()+1);
               if ((tmpCell == null) || (!equals(resCell, tmpCell,  CellParam.WIDTH))){
                   next = false;

               }
               else if (isLabel(tmpCell)) {
                   next = false;
               }
               else if (tmpCell.getStyle().getTopBorder().getType() != BorderType.NONE)
                   next = false;

               if (next){
                   resCell = tmpCell;
                   ul = resCell.getRt();
               }

           }
           else break;


        }while (next);

        if (next == false){
            //Cells may be merged
            int cellTop = curCell.getRb()+1;
            while(ul > cellTop){
                tmpCell = getCellByCoord(curCell.getCl(), cellTop);
                curCell = curCell.merge(tmpCell);
                cellTop = tmpCell.getRb()+1;

            };
        }

        return curCell;
    }
    private CCell checkRight(CCell curCell){
        return checkRight(curCell, -1);
    }
    private CCell checkRight(CCell curCell, int rightBorder){
        CCell resCell = curCell, tmpCell;
        CStyle cellStyle;
        CBorder cB;
        boolean next;
        do{
            next = true;
            if (resCell.getCr()+1 < hR){
                if (resCell.getStyle().getRightBorder().getType() != BorderType.NONE)
                    break;
                tmpCell = getCellByCoord(resCell.getCr()+1, resCell.getRt());
                if (isLabel(tmpCell) || (!equals(resCell, tmpCell,  CellParam.HEIGHT))) break;
                if ((tmpCell == null) || (tmpCell.getStyle().getLeftBorder().getType() != BorderType.NONE))
                    break;
                else {
                    cB = tmpCell.getStyle().getRightBorder();
                    resCell = resCell.merge(tmpCell);
                    cellStyle = resCell.getStyle();
                    cellStyle.setRightBorder(cB);
                    resCell.setStyle(cellStyle);


                }

                    cellTransofrm(resCell, new Block(resCell)); //Was added for test
            }
            else next = false;
        }while (next);

        return  resCell;
    }

    private int analyse2BottomLine(CCell cCell, int bottomBorder){
        //Looking for bottom border of col
        CCell curCell = cCell, tmpCell;

        int border = curCell.getRb();
        do{
            if (curCell.getRb() > bottomBorder)
                return -1;
            //Bottom border exists
            if ((curCell.getStyle().getBottomBorder().getType() != BorderType.NONE) || (curCell.getRb() == bottomBorder))
                return curCell.getRb();
            tmpCell = getCellByCoord(curCell.getCr(), curCell.getRb()+1);
            //Lower cell upper border exists
            if (tmpCell.getStyle().getTopBorder().getType() != BorderType.NONE)
                return curCell.getRb();
            curCell = tmpCell;
            //Check text



        }while(true);
    }


    private CCell mergeToLeft(CCell cCell){
        CCell leftCell;
        //Check the text existance in the cell
        if (isLabel(cCell) || (cCell.getStyle().getLeftBorder().getType() != BorderType.NONE))
            return cCell;
        //The cell does not contain any text
        leftCell = getLeftCell(cCell);
        if (leftCell == null)
            return  cCell;
        if (leftCell.getStyle().getRightBorder().getType() != BorderType.NONE)
            return cCell;
        if ((cCell.getRt() == leftCell.getRt()) && (cCell.getRb() == leftCell.getRb()))
            cCell = leftCell.merge(cCell);
        return cCell;

    }


}
