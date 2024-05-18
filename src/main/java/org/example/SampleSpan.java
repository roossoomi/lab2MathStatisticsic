package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class SampleSpan {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public SampleSpan(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> getSampleSpan(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("размах", StatUtils.max(dataarray)-StatUtils.min(dataarray));
        return dictionary;
    }
}
