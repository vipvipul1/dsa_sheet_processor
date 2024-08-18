package com.vip.dsa;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

public class DsaSheetProcessorApp {
    public static void main(String[] args) throws IOException {
        String path = "/Users/vipulkumar/Library/CloudStorage/OneDrive-Personal/Documents/Scaler_Cloud/DSA_Bank.xlsx";
        CombineSheets combineSheets = new CombineSheets();
        combineSheets.mergeAllSheets(path);
    }
}
