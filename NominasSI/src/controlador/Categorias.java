package controlador;
// Generated 14-mar-2023 11:36:10 by Hibernate Tools 4.3.1



/**
 * Categorias generated by hbm2java
 */
public class Categorias  implements java.io.Serializable {


     private int idCategoria;
     private String nombreCategoria;
     private double salarioBaseCategoria;
     private double complementoCategoria;

    public Categorias() {
    }

    public Categorias(int idCategoria, String nombreCategoria, double salarioBaseCategoria, double complementoCategoria) {
       this.idCategoria = idCategoria;
       this.nombreCategoria = nombreCategoria;
       this.salarioBaseCategoria = salarioBaseCategoria;
       this.complementoCategoria = complementoCategoria;
    }
   
    public int getIdCategoria() {
        return this.idCategoria;
    }
    
    public void setIdCategoria(int idCategoria) {
        this.idCategoria = idCategoria;
    }
    public String getNombreCategoria() {
        return this.nombreCategoria;
    }
    
    public void setNombreCategoria(String nombreCategoria) {
        this.nombreCategoria = nombreCategoria;
    }
    public double getSalarioBaseCategoria() {
        return this.salarioBaseCategoria;
    }
    
    public void setSalarioBaseCategoria(double salarioBaseCategoria) {
        this.salarioBaseCategoria = salarioBaseCategoria;
    }
    public double getComplementoCategoria() {
        return this.complementoCategoria;
    }
    
    public void setComplementoCategoria(double complementoCategoria) {
        this.complementoCategoria = complementoCategoria;
    }




}

