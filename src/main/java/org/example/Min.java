package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class Min {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public Min(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> getMin(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("минимум", StatUtils.min(dataarray));
        return dictionary;
    }
}
