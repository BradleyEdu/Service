package proyecto_serviciosocial;

import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class prueba {
    
    
    public static void consulta(){
        conexion conect = new conexion();
        Connection con = conect.getConnection();
        String v1 = "", v2 = "";
        int i = 0;
            
    
        try {
            System.out.println("---------------Aqui inicia--------------");
            String sql = "select filiacion, nombre from personal LIMIT 5";
            
            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);
            
            while(rs.next()){
                v1 = rs.getString(1);
                v2 = rs.getString(2);
                System.out.println(i + ": Filicacion: " + v1 + ", Nombre: " + v2);
            }
            
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("Error de consulta");
        }
    }
    
    public static void main(String [] args){
        
        consulta();
        
    }
   

}
    