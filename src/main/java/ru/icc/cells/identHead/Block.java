package ru.icc.cells.identHead;

import ru.icc.cells.ssdc.model.CCell;

public class Block {
    private int top, bottom, left, right;
    private boolean isLabel = false;

    Block(CCell cCell){
        top = cCell.getRt();
        bottom = cCell.getRb();
        left = cCell.getCl();
        right = cCell.getCr();
    }

    Block(int top, int bottom, int left, int right){
        this.top = top;
        this.bottom = bottom;
        this.left = left;
        this.right = right;
    }

    public void setTop(int top) {
        this.top = top;
    }

    public void setBottom(int bottom) {
        this.bottom = bottom;
    }

    public void setLeft(int left) {
        this.left = left;
    }

    public void setRight(int right) {
        this.right = right;
    }

    public void setLabel(boolean label) {
        isLabel = label;
    }
    public int getTop(){
        return top;
    }

    public int getBottom() {
        return bottom;
    }

    public int getLeft() {
        return left;
    }

    public int getRight() {
        return right;
    }

    public boolean getIsLabel(){return isLabel; }

    public void increaseBlockSize(CCell cCell){
        if (cCell == null) return;
        if (cCell.getRb() > bottom) bottom = cCell.getRb();
        if (cCell.getCr() > right) right = cCell.getCr() +1;

    }

    public boolean concatBlock(Block block){
        if ((top == block.top) || (bottom == block.bottom)){
            //Horisintal block concatination
            if (((left <=block.left) && (right+1 >= block.left)) || (left >= block.left) && (left <= block.right+1)){
                if (left > block.left)
                    left = block.left;
                if (right< block.right)
                    right = block.right;
                return true;
            }
            if ((left == block.left) || (right == block.right)){
                //Vertical concatination
                if(((block.top <= top) && (block.bottom+1 >= top)) || ((block.top<=bottom) && (block.bottom+1>=bottom))){
                    if (top > block.top)
                        top = block.top;
                    if (bottom < block.bottom)
                        bottom = block.bottom;
                    return true;
                }
            }
        }
        return false;
    }

}
