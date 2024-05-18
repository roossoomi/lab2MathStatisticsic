package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class Variance {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public Variance(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public  HashMap<String, Double> getVariance(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("дисперсия", StatUtils.variance(dataarray));
        return dictionary;
    }
}
