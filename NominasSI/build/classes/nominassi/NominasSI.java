/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package nominassi;

import controlador.Nomina;
import controlador.Trabajador;
import java.io.IOException;

import nominassi.ExcelManager;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.List;
import java.util.Scanner;
//import org.hibernate.Session;
//import org.hibernate.SessionFactory;
import util.HibernateUtil;

/**
 *
 * @author valen
 */
public class NominasSI {

    /**
     * @param args the command line arguments
     */
    
    
    private static final String URL="jdbc:mysql://localhost:3306/nominas";
    private static final String USUARIO = "root";
    public static final String PASSWORD = "1234";
    
    //coonexion
    
    static Connection conexion = null;
    
    
    public static void main(String[] args) throws SQLException, IOException {
        
        /*
        SessionFactory sf = HibernateUtil.getSessionFactory();
        Session session = sf.openSession();
        
        
        System.out.println("Introducir el CIF del trabajador:");
        Scanner scan = new Scanner(System.in);
        String cif = scan.nextLine();
        Trabajador trabajadorDAO = new Trabajador();
        trabajadorDAO.setConector(session);
        Trabajador trabajador = trabajadorDAO.encontrarPorCif(cif);
        if (trabajador == null) {
            System.out.println("No hemos encontrado al trabajador en nuestro sistema");
        } else {
            //datos Trabajador
            System.out.println("Nombre trabajador: " + trabajador.getNombre());
        
    }
    
    */
    
    //Ejercicio 2
    
        String rutaExcel = "C:/Users/Torre/Documents/GitHub/Practica_SI/NominasSI/src/resources/SistemasInformacionII.xlsx";
        
        ExcelManager resolverEjercicio = new ExcelManager();
        

         try {
             /*
             List<String> listaNombres = resolverEjercicio.obtenerFila(rutaExcel, "09548150E"); //tener en cuenta que las hojas comienzan en 0
             for (String celda : listaNombres) {
                    System.out.println(celda);
              } */
             
             //resolverEjercicio.modificarDatos(rutaExcel, 1, "Cocinero", "lcoo");
             //resolverEjercicio.procesarDNI(rutaExcel);
             resolverEjercicio.procesarDNI(rutaExcel);
           
             
            
            
        } catch (Exception ex) {
           ex.printStackTrace();
        
        }
    }

}