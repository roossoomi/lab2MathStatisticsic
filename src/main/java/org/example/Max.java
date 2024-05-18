package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class Max {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public Max(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> getMax(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("максимум", StatUtils.max(dataarray));
        return dictionary;
    }
}
