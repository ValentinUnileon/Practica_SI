package controlador;
// Generated 14-mar-2023 11:36:10 by Hibernate Tools 4.3.1



/**
 * Empresas generated by hbm2java
 */
public class Empresas  implements java.io.Serializable {


     private int idEmpresa;
     private String nombre;
     private String cif;

    public Empresas() {
    }

    public Empresas(int idEmpresa, String nombre, String cif) {
       this.idEmpresa = idEmpresa;
       this.nombre = nombre;
       this.cif = cif;
    }
   
    public int getIdEmpresa() {
        return this.idEmpresa;
    }
    
    public void setIdEmpresa(int idEmpresa) {
        this.idEmpresa = idEmpresa;
    }
    public String getNombre() {
        return this.nombre;
    }
    
    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    public String getCif() {
        return this.cif;
    }
    
    public void setCif(String cif) {
        this.cif = cif;
    }




}


