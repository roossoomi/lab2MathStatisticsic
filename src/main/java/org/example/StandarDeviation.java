package org.example;

import org.apache.commons.math3.stat.StatUtils;

import java.util.HashMap;

public class StandarDeviation {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public StandarDeviation(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> getStandarDeviation(HashMap<String, Double> dictionary,double[] dataarray) {
        double geometricMean = StatUtils.variance(dataarray);
        // Вычисление среднего геометрического
        double variace  = Math.pow(geometricMean, 0.5);
        dictionary.put("стандартного отклонения", variace);
        return dictionary;
    }
}
