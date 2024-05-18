package org.example;

import java.util.HashMap;

public class Calculation {
    private  double[] dataarray;
    private HashMap<String, Double> dictionary;
    public Calculation(HashMap<String, Double> dictionary,double[] dataarray){
        this.dictionary=dictionary;
        this.dataarray=dataarray;
    }
    public HashMap<String, Double> makeCalculation(HashMap<String, Double> dictionary,double[] dataarray) {
        NumberOfElements numberOfElementsexample = new NumberOfElements(dictionary,dataarray);
        ArithmeticMean arithmeticMeanexample = new ArithmeticMean(dictionary,dataarray);
        coefficientOfVariation coefficientOfVariationexample = new coefficientOfVariation(dictionary,dataarray);
        GeometricMeanCalculator geometricMeanCalculatorexample = new GeometricMeanCalculator(dictionary,dataarray);
        Max maxexample = new Max(dictionary,dataarray);
        Min minexample = new Min(dictionary,dataarray);
        SampleSpan sampleSpanexample = new SampleSpan(dictionary,dataarray);
        StandarDeviation standarDeviationexample = new StandarDeviation(dictionary,dataarray);
        Variance varianceexample = new Variance(dictionary,dataarray);
        arithmeticMeanexample.getArithmeticMean(dictionary,dataarray);
        coefficientOfVariationexample.getCoefficientOfVariation(dictionary,dataarray);
        geometricMeanCalculatorexample.getGeometricMeanCalculator(dictionary,dataarray);
        maxexample.getMax(dictionary,dataarray);
        minexample.getMin(dictionary,dataarray);
        numberOfElementsexample.getNumberOfElements(dictionary,dataarray);
        sampleSpanexample.getSampleSpan(dictionary,dataarray);
        standarDeviationexample.getStandarDeviation(dictionary,dataarray);
        varianceexample.getVariance(dictionary,dataarray);
        return dictionary;
    }
}
