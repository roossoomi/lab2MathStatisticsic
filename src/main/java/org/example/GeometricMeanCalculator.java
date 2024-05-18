package org.example;

import java.util.HashMap;

public class GeometricMeanCalculator {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public GeometricMeanCalculator(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }

    public HashMap<String, Double> getGeometricMeanCalculator(HashMap<String, Double> dictionary,double[] dataarray) {
        double product = 1.0;
        for (Double str : dataarray) {
            product *= str;
        }

        // Вычисление среднего геометрического
        double geometricMean = Math.pow(product, 1.0 / dataarray.length);
        dictionary.put("Среднее геометрическое", geometricMean);
        return dictionary;
    }
}
