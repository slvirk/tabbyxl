package ru.icc.cells.identHead;

public class CellCoordNode {
    private int cl;
    private int cr;
    private int ct;
    private int cb;

    CellCoordNode(int l, int r, int t, int b){
        cl = l;
        cr = r;
        ct = t;
        cb = b;

    }

    public int getCl() {
        return cl;
    }

    public int getCr() {
        return cr;
    }

    public int getCt() {
        return ct;
    }

    public int getCb(){
        return cb;
    }

    public void setCl(int cl) {
        this.cl = cl;
    }

    public void setCr(int cr) {
        this.cr = cr;
    }

    public void setCt(int ct) {
        this.ct = ct;
    }

    public void setCb(int cb) {
        this.cb = cb;
    }

}
