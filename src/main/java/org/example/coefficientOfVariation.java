package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class coefficientOfVariation {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;

    public coefficientOfVariation(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> getCoefficientOfVariation(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("коэф вариации", Math.sqrt(StatUtils.variance(dataarray))/StatUtils.mean(dataarray));
        return dictionary;
    }
}
