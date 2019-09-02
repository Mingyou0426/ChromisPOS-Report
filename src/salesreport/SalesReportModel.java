/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package salesreport;

/**
 *
 * @author Jincowboy
 */
public class SalesReportModel {
    String dateval;
    String quantitysold;
    String costvalue;
    String salesvalue;
    String profit;

    public SalesReportModel(String dateval, String quantitysold, String costvalue, String salesvalue, String profit) {
        this.dateval = dateval;
        this.quantitysold = quantitysold;
        this.costvalue = costvalue;
        this.salesvalue = salesvalue;
        this.profit = profit;
    }

    public String getCostvalue() {
        return costvalue;
    }

    public void setCostvalue(String costvalue) {
        this.costvalue = costvalue;
    }

    public String getDateval() {
        return dateval;
    }

    public void setDateval(String dateval) {
        this.dateval = dateval;
    }

    public String getProfit() {
        return profit;
    }

    public void setProfit(String profit) {
        this.profit = profit;
    }

    public String getQuantitysold() {
        return quantitysold;
    }

    public void setQuantitysold(String quantitysold) {
        this.quantitysold = quantitysold;
    }

    public String getSalesvalue() {
        return salesvalue;
    }

    public void setSalesvalue(String salesvalue) {
        this.salesvalue = salesvalue;
    }

    
}
