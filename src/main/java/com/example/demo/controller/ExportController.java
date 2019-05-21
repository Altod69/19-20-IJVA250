package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.*;

import org.apache.poi.ss.formula.functions.FactDouble;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;
    
    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");
        
        for(Client client : allClients) {
        	writer.println(
        			client.getId() + ";"
        		 + client.getNom() + ";" 
        		 + client.getPrenom() + ";" 
        		 + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY"))
        		 );
        }
        
        
    }
    
    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vmd.excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        
        List<Client> allClients = clientService.findAllClients();
        
        
    	Workbook workbook = new XSSFWorkbook();
    	Sheet sheet = workbook.createSheet("clients");
    	Row headerRow = sheet.createRow(0);
    	
    	Cell cellId = headerRow.createCell(0);
    	cellId.setCellValue("Id");
    	
    	Cell cellNom = headerRow.createCell(1);
    	cellNom.setCellValue("Nom");
    	
    	Cell cellPrenom = headerRow.createCell(2);
    	cellPrenom.setCellValue("Prénom");
    	
    	Cell cellDateNaissance = headerRow.createCell(3);
    	cellDateNaissance.setCellValue("Date de naissance");
    	
    	int i=1;
    	
        for(Client client : allClients) {
        	Row completeRow = sheet.createRow(i);
        	
        	Cell cellCompleteId = completeRow.createCell(0);
        	cellCompleteId.setCellValue(client.getId());
        	
        	Cell cellCompleteNom = completeRow.createCell(1);
        	cellCompleteNom.setCellValue(client.getNom());
        	
        	Cell cellCompletePrenom = completeRow.createCell(2);
        	cellCompletePrenom.setCellValue(client.getPrenom());
        	
        	Cell cellCompleteDateNaissance = completeRow.createCell(3);
        	cellCompleteDateNaissance.setCellValue(client.getDateNaissance().toString());
        	
        	i++;
        }
        
    	workbook.write(response.getOutputStream());
    	workbook.close();    
    }
    
    @RequestMapping(value = "/clients/{{id}}/factures/xlsx", method = RequestMethod.GET)
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response, @PathVariable Long clientId) throws IOException {
        response.setContentType("application/vmd.excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        
        Workbook workbook = new XSSFWorkbook();
    	Sheet sheet = workbook.createSheet("factures");
    	Row headerRow = sheet.createRow(0);
    	
    	Cell cellId = headerRow.createCell(0);
    	cellId.setCellValue("Id");
    	
    	Cell cellTotal = headerRow.createCell(1);
    	cellTotal.setCellValue("Total");
        
    	List<Facture> allFactures = factureService.findByClientId(clientId);
        
    	int i = 1;
        for(Facture facture : allFactures) {
        	Row completeRow = sheet.createRow(i);
        	
        	Cell cellCompleteId = completeRow.createCell(0);
        	cellCompleteId.setCellValue(facture.getId());
        	
        	Cell cellCompleteTotal = completeRow.createCell(1);
        	cellCompleteTotal.setCellValue(facture.getTotal());
        	
        	i++;
        }
        
    	workbook.write(response.getOutputStream());
    	workbook.close();    
    }
}
