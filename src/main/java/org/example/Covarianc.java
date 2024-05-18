package org.example;

import org.apache.commons.math3.stat.correlation.Covariance;

import java.util.HashMap;

public class Covarianc {

    private  double[][] dataarray;
    public Covarianc(double[][] dataarray){
        this.dataarray=dataarray;
    }


    public static HashMap<String, Double> getCovariance(HashMap<String, double[]> mapData) {
        HashMap<String, Double> covariationList = new HashMap<>();
        Covariance covariance = new Covariance();
        for (String key1 : mapData.keySet()){
            for (String key2 : mapData.keySet()){
                covariationList.put(key1+key2,covariance.covariance(mapData.get(key1), mapData.get(key2)));
            }
        }
        return covariationList;
    }
}