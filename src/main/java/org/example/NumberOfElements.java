package org.example;

import java.util.HashMap;

public class NumberOfElements {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public NumberOfElements(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public  HashMap<String, Double> getNumberOfElements(HashMap<String, Double> dictionary,double[] dataarray) {
        dictionary.put("кол-во элементов", Double.valueOf(dataarray.length));
        return dictionary;
    }
}
