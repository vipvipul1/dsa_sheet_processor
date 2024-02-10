package com.vip.dsa;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

public class DsaSheetProcessorApp {
    public static void main(String[] args) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        String path = reader.readLine();
        CombineSheets combineSheets = new CombineSheets();
        combineSheets.mergeAllSheets(path);
    }
}
