package com.example.demo.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.example.demo.entity.PostData;

@RestController
public class apiCall {

    @PostMapping("/calling_api")
    public List<Map<String, String>> calling(@RequestBody PostData data) {
        String url = "https://iabpiz-test.fa.ocs.oraclecloud.com:443/xmlpserver/services/ExternalReportWSSService";
        String srId = data.getsr_id();
        String reportPath = data.getpath();

        // SOAP Request Body
        String xmlData = String.format(""" 
            <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope" xmlns:pub="http://xmlns.oracle.com/oxp/service/PublicReportService">
                <soap:Body>
                    <pub:runReport>
                        <pub:reportRequest>
                            <pub:attributeFormat>xlsx</pub:attributeFormat>
                            <pub:parameterNameValues>
                                <pub:item>
                                    <pub:name>SrNumber</pub:name>
                                    <pub:values>
                                        <pub:item>%s</pub:item>
                                    </pub:values>
                                </pub:item>
                            </pub:parameterNameValues>
                            <pub:reportAbsolutePath>%s</pub:reportAbsolutePath>
                            <pub:sizeOfDataChunkDownload>-1</pub:sizeOfDataChunkDownload>
                        </pub:reportRequest>
                    </pub:runReport>
                </soap:Body>
            </soap:Envelope>
            """, srId, reportPath);

        // Authorization token
        String authToken = "Basic YWxvay5rdW1hcjpmdXNpb25AMTIz";

        try {
            // Build the HTTP Request
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Content-Type", "application/soap+xml")
                    .header("Authorization", authToken)
                    .POST(HttpRequest.BodyPublishers.ofString(xmlData))
                    .build();

            // Send the HTTP Request
            HttpClient client = HttpClient.newHttpClient();
            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

            System.out.println("Status Code: " + response.statusCode());
            System.out.println("Response Body: " + response.body());

            if (response.statusCode() != 200) {
                throw new IOException("Error: " + response.statusCode());
            }

            var responseBody = response.body();
            String regex = "<ns2:reportBytes>(.*?)</ns2:reportBytes>";

            // Compile the regular expression
            Pattern pattern = Pattern.compile(regex, Pattern.DOTALL);
            Matcher matcher = pattern.matcher(responseBody);
            String result = "";
            if (matcher.find()) {
                // Extract the string between the tags
                result = matcher.group(1);
            } else {
                System.out.println("No match found");
            }

            byte[] decodedBytes = Base64.getDecoder().decode(result);
            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(decodedBytes);

            // Read the Excel file using Apache POI
            Workbook workbook = new XSSFWorkbook(byteArrayInputStream);
            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet

            List<Map<String, String>> dataList = new ArrayList<>();
            Row headerRow = sheet.getRow(0); // Get header row

            for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);

                // Check if the row is null
                if (row == null) {
                    continue;  // Skip processing this row if it's null
                }

                Map<String, String> rowMap = new HashMap<>();

                for (int j = 0; j < headerRow.getPhysicalNumberOfCells(); j++) {
                    String cellValue = row.getCell(j) != null ? row.getCell(j).toString() : "";
                    rowMap.put(headerRow.getCell(j).toString(), cellValue);
                }

                dataList.add(rowMap);
            }

            // Close workbook
            workbook.close();

            return dataList; // Return as JSON directly

        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
            throw new RuntimeException("An error occurred: " + e.getMessage());
        }
    }
}
