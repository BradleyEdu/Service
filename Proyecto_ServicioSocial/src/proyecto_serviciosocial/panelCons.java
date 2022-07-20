package proyecto_serviciosocial;

import static java.awt.Color.black;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bradl
 */
public class panelCons extends javax.swing.JPanel {

    conexion conect = new conexion();
    Connection con = conect.getConnection();
    DefaultTableModel modelo = new DefaultTableModel();
    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

    public panelCons() {
        initComponents();
        modelo.addColumn("Del");
        modelo.addColumn("Al");
        modelo.addColumn("Código");
        modelo.addColumn("Puesto");
        modelo.addColumn("Nivel");
        modelo.addColumn("Zona");
        modelo.addColumn("Sueldo");
        modelo.addColumn("Quinquenio");
        modelo.addColumn("Otras");
        modelo.addColumn("Total");
        modelo.addColumn("Motivo");
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        txtFili = new javax.swing.JTextField();
        txtCurp = new javax.swing.JTextField();
        txtRfc = new javax.swing.JTextField();
        txtHomo = new javax.swing.JTextField();
        btnBusc = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        tablaPuestos = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        lblNom = new javax.swing.JLabel();
        lblPaterno = new javax.swing.JLabel();
        btnLimpi = new javax.swing.JButton();
        btnImprimir = new javax.swing.JButton();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel10 = new javax.swing.JLabel();
        lblFechaIng = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        lblFechaBaja = new javax.swing.JLabel();
        lblMotivos = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        lblMaterno = new javax.swing.JLabel();
        lblMaterno1 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel18 = new javax.swing.JLabel();
        jSeparator2 = new javax.swing.JSeparator();
        lblCalle = new javax.swing.JLabel();
        lblNumExt = new javax.swing.JLabel();
        jLabel19 = new javax.swing.JLabel();
        jLabel20 = new javax.swing.JLabel();
        lblColonia = new javax.swing.JLabel();
        jLabel21 = new javax.swing.JLabel();
        lblCodP = new javax.swing.JLabel();
        lblCiudad = new javax.swing.JLabel();
        jLabel22 = new javax.swing.JLabel();
        jLabel23 = new javax.swing.JLabel();
        lblEstado = new javax.swing.JLabel();
        jLabel24 = new javax.swing.JLabel();
        jSeparator3 = new javax.swing.JSeparator();
        jButton1 = new javax.swing.JButton();

        setBackground(new java.awt.Color(255, 255, 255));
        setPreferredSize(new java.awt.Dimension(500, 500));

        txtFili.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCurp.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtRfc.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtHomo.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        txtHomo.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtHomoActionPerformed(evt);
            }
        });

        btnBusc.setText("Buscar");
        btnBusc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnBuscActionPerformed(evt);
            }
        });

        tablaPuestos.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane1.setViewportView(tablaPuestos);

        jLabel1.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel1.setText("Datos personales");

        jLabel2.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel2.setText("Nombres");

        jLabel3.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel3.setText("Apellido Paterno");

        lblNom.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        lblPaterno.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        btnLimpi.setText("Limpiar");

        btnImprimir.setText("Imprimir");
        btnImprimir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnImprimirActionPerformed(evt);
            }
        });

        jLabel6.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel6.setText("Filiacion");

        jLabel7.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel7.setText("Curp");

        jLabel8.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel8.setText("RFC");

        jLabel9.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel9.setText("Homoclave");

        jSeparator1.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator1.setForeground(new java.awt.Color(0, 0, 0));

        jLabel10.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel10.setText("Fecha de ingreso");

        lblFechaIng.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel12.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel12.setText("Motivos de la baja");

        jLabel13.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel13.setText("Fecha de baja");

        lblFechaBaja.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        lblMotivos.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel4.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel4.setText("Apellido Materno");

        lblMaterno.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        lblMaterno1.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N

        jLabel5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel5.setText("Domicilio");

        jLabel18.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel18.setText("Calle");

        jSeparator2.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator2.setForeground(new java.awt.Color(0, 0, 0));

        lblCalle.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        lblNumExt.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel19.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel19.setText("Numero Exterior");

        jLabel20.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel20.setText("Colonia o Fraccionamiento");

        lblColonia.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel21.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel21.setText("Código postal");

        lblCodP.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        lblCiudad.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel22.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel22.setText("Ciudad");

        jLabel23.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel23.setText("Estado");

        lblEstado.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N

        jLabel24.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel24.setText("Periodo de aportaciones al fondo del ISSTE");

        jSeparator3.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator3.setForeground(new java.awt.Color(0, 0, 0));

        jButton1.setBackground(new java.awt.Color(255, 255, 255));
        jButton1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/proyecto_serviciosocial/add_box_FILL0_wght400_GRAD0_opsz48.png"))); // NOI18N

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(40, 40, 40)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel18)
                                            .addComponent(lblCalle, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(27, 27, 27)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(jLabel19)
                                            .addComponent(jLabel22)
                                            .addComponent(lblCiudad, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(lblNumExt, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(38, 38, 38)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel20)
                                            .addComponent(lblColonia, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jLabel23)
                                            .addComponent(lblEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 169, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(180, 180, 180))
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel5)
                                            .addComponent(jLabel1)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(jLabel2)
                                                    .addComponent(lblNom, javax.swing.GroupLayout.PREFERRED_SIZE, 261, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                .addGap(47, 47, 47)
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(jLabel3)
                                                    .addComponent(lblPaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 174, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                .addGap(56, 56, 56)
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(jLabel4)
                                                    .addComponent(lblMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 174, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                        .addGap(0, 0, Short.MAX_VALUE)))
                                .addGap(28, 28, 28))
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 746, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 740, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jSeparator3, javax.swing.GroupLayout.PREFERRED_SIZE, 742, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel10)
                                    .addComponent(jLabel24)
                                    .addComponent(jLabel8)
                                    .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 212, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtFili, javax.swing.GroupLayout.PREFERRED_SIZE, 202, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel6))
                                        .addGap(78, 78, 78)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel7)
                                            .addGroup(layout.createSequentialGroup()
                                                .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 213, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addGap(104, 104, 104)
                                                .addComponent(btnBusc)
                                                .addGap(36, 36, 36)
                                                .addComponent(btnLimpi))
                                            .addComponent(jLabel9)
                                            .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 283, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                    .addComponent(jLabel21)
                                    .addComponent(lblCodP, javax.swing.GroupLayout.PREFERRED_SIZE, 174, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                        .addGroup(layout.createSequentialGroup()
                                            .addComponent(jLabel12)
                                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(lblMotivos, javax.swing.GroupLayout.PREFERRED_SIZE, 388, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                                            .addComponent(lblFechaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 256, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addGap(55, 55, 55)
                                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                .addComponent(jLabel13)
                                                .addComponent(lblFechaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 282, javax.swing.GroupLayout.PREFERRED_SIZE)))))
                                .addGap(146, 146, 146))))
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 949, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(381, 381, 381)
                        .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(58, 58, 58)
                        .addComponent(btnImprimir)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 396, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(lblMaterno1, javax.swing.GroupLayout.PREFERRED_SIZE, 174, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(422, 422, 422)
                .addComponent(lblMaterno1, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(jLabel7))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtFili, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnBusc)
                    .addComponent(btnLimpi))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel8)
                    .addComponent(jLabel9))
                .addGap(5, 5, 5)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(34, 34, 34)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(lblNom, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel4)
                            .addComponent(jLabel3))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(lblPaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(43, 43, 43)
                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel18)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(lblCalle, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel19)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(lblNumExt, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGap(18, 18, 18)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel22)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(lblCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel21)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(lblCodP, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel20)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(lblColonia, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(77, 77, 77))
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel23)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(lblEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(26, 26, 26)
                        .addComponent(jLabel24, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jSeparator3, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel10)
                            .addComponent(jLabel13))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(lblFechaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(lblFechaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel12)
                        .addGap(12, 12, 12))
                    .addComponent(lblMotivos, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 45, Short.MAX_VALUE)
                .addGap(23, 23, 23)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jButton1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addComponent(btnImprimir)
                        .addGap(8, 8, 8)))
                .addContainerGap())
        );

        getAccessibleContext().setAccessibleParent(this);
    }// </editor-fold>//GEN-END:initComponents

    private void btnImprimirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnImprimirActionPerformed
        try {
            llenarNombre();
        } catch (IOException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_btnImprimirActionPerformed

    private void txtHomoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtHomoActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtHomoActionPerformed

    private void btnBuscActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnBuscActionPerformed
        try {
            consultarPersonal();
        } catch (IOException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_btnBuscActionPerformed

    public String sacarFechaBaja(String filia){
        try {
            String sql = "SELECT personal.baja FROM personal_sueldo, personal WHERE personal_sueldo.motivo != '"+null+"' AND "
                    + "personal_sueldo.filiacion = '" + filia + "'";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                Date fechaBaja = rs.getDate(1);
                if(rs.getDate(1) != null){
                    return fechaBaja.toString();
                }  
            }

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta de la fecha", JOptionPane.ERROR_MESSAGE);
        }
        return "";
    }
    
    public void consultarPersonal() throws IOException {

        try {
            String sql = "select curp, rfc, homoclave, apellido_paterno, apellido_materno"
                    + ", nombre, calle, numero, colonia, codigo_postal, ciudad, estado"
                    + ", ingreso, baja, motivo_baja from personal where filiacion = '" + txtFili.getText()+"'" ;

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                txtCurp.setText(rs.getString(1));
                txtRfc.setText(rs.getString(2));
                txtHomo.setText(rs.getString(3));
                lblPaterno.setText(rs.getString(4));
                lblMaterno.setText(rs.getString(5));
                lblNom.setText(rs.getString(6));
                lblCalle.setText(rs.getString(7));
                lblNumExt.setText(rs.getString(8));
                lblColonia.setText(rs.getString(9));
                lblCodP.setText(rs.getString(10));
                lblCiudad.setText(rs.getString(11));
                lblEstado.setText(rs.getString(12));
                Date fechaIngreso = rs.getDate(13);
                lblFechaIng.setText(formatter.format(fechaIngreso));
                if (rs.getDate(14) != null) {
                    Date fechaBaja = rs.getDate(14);
                    lblFechaBaja.setText(formatter.format(fechaBaja));
                }
                lblMotivos.setText(rs.getString(15));
            }

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta", JOptionPane.ERROR_MESSAGE);
        }

        consultarPersonal_Sueldo();
    }

    public void consultarPersonal_Sueldo() throws IOException {
        Object columnas[] = new Object[11];
        
        try {
            String sql = "SELECT del, al, codigo, puesto, nivel, zona, sueldo, quinquenio, otras, total, motivo FROM personal_sueldo WHERE filiacion = '" + txtFili.getText() + "'"
                    /*+ " AND del <= '" + sacarFechaBaja(txtFili.getText()) + "'"*/
                    + " ORDER BY del ASC";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);
            int conta = 29;
            while (rs.next()) {
                Date fechaDel = rs.getDate(1);
                Date fechaAl = rs.getDate(2);
                columnas[0] = formatter.format(fechaDel);
                columnas[1] = formatter.format(fechaAl);
                columnas[2] = rs.getString(3);
                columnas[3] = rs.getString(4);
                columnas[4] = rs.getString(5);
                columnas[5] = rs.getString(6);
                columnas[6] = rs.getString(7);
                columnas[7] = rs.getString(8);
                columnas[8] = rs.getString(9);
                columnas[9] = rs.getString(10);
                columnas[10] = rs.getString(11);
                modelo.addRow(columnas);
                llenarTabla(columnas, conta);
                conta++;
            }
            tablaPuestos.setModel(modelo);

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    public void llenarNombre() throws IOException{
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow fila = sheet.getRow(6);

            if(fila == null){
                fila = sheet.createRow(6);
            }

            XSSFCell nombre = fila.createCell(4);
            XSSFCell apellido_pat = fila.createCell(0);
            XSSFCell apellido_mat = fila.createCell(2);
            XSSFCell curp = fila.createCell(8);
            XSSFCell rfc = fila.createCell(6);
            
            if(nombre == null){
                nombre = fila.createCell(4);
            }
            
            if(apellido_pat == null){
                apellido_pat = fila.createCell(0);
            }
            
            if(apellido_mat == null){
                apellido_pat = fila.createCell(2);
            }
                        
            if(curp == null){
                curp = fila.createCell(8);
            }
            
            if(rfc == null){
                rfc = fila.createCell(6);
            }
            


            CellStyle cellStyle = fila.getSheet().getWorkbook().createCellStyle();
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
            
            Font cellFont = wb.createFont();
            cellFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            cellStyle.setFont(cellFont);
        
            nombre.setCellStyle(cellStyle);
            apellido_pat.setCellStyle(cellStyle);
            apellido_mat.setCellStyle(cellStyle);
            curp.setCellStyle(cellStyle);
            rfc.setCellStyle(cellStyle);
            
            nombre.setCellValue(lblNom.getText());
            apellido_pat.setCellValue(lblPaterno.getText());
            apellido_mat.setCellValue(lblMaterno.getText());
            curp.setCellValue(txtCurp.getText());
            rfc.setCellValue(txtRfc.getText());
            

            file.close();

            FileOutputStream output = new FileOutputStream("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            wb.write(output);
            output.close();
            
            File file1= new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            Desktop.getDesktop().open(file1);
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void llenarTabla(Object[] datos, int cont) throws IOException{
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow fila = sheet.getRow(cont);
            
            
            if(fila == null){
                fila = sheet.createRow(cont);
            }

            XSSFCell celda1 = fila.createCell(0); //Motivo
            XSSFCell celda2 = fila.createCell(1); //del
            XSSFCell celda3 = fila.createCell(2); //al
            XSSFCell celda4 = fila.createCell(3); //nombre
            XSSFCell celda5 = fila.createCell(4); //puesto
            XSSFCell celda6 = fila.createCell(5); //ramo registrado
            XSSFCell celda7 = fila.createCell(6); //pagaduria registrada
            XSSFCell celda8 = fila.createCell(7); //saldo basico
            XSSFCell celda9 = fila.createCell(8); //quinquenio
            XSSFCell celda10 = fila.createCell(9); //otras percepciones
            XSSFCell celda11 = fila.createCell(10); //total
            
            if(celda1 == null){
                celda1 = fila.createCell(0);
            }
            
            if(celda2 == null){
                celda2 = fila.createCell(1);
            }
            
            if(celda3 == null){
                celda3 = fila.createCell(2);
            }
            
            if(celda4 == null){
                celda4 = fila.createCell(3);
            }
            
            if(celda5 == null){
                celda5 = fila.createCell(4);
            }
            
            if(celda6 == null){
                celda6 = fila.createCell(5);
            }
            
            if(celda7 == null){
                celda7 = fila.createCell(6);
            }
            
            if(celda8 == null){
                celda8 = fila.createCell(7);
            }
            
            if(celda9 == null){
                celda9 = fila.createCell(8);
            }
            
            if(celda10 == null){
                celda10 = fila.createCell(9);
            }
            
            if(celda11 == null){
                celda11 = fila.createCell(10);
            }
            
            for(int i = 0; i<datos.length;i++){
                if (datos[i] == null){
                    datos[i] = "";
                }
            }
            
            celda1.setCellValue(datos[0].toString());
            celda2.setCellValue(datos[1].toString());          
            celda3.setCellValue(datos[2].toString());
            celda4.setCellValue(datos[3].toString());
            celda5.setCellValue(datos[4].toString());
            celda6.setCellValue(datos[5].toString());
            celda7.setCellValue(datos[6].toString());
            celda8.setCellValue(datos[7].toString());
            celda9.setCellValue(datos[8].toString());
            celda10.setCellValue(datos[9].toString());
            celda11.setCellValue(datos[10].toString());

            file.close();

            FileOutputStream output = new FileOutputStream("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            wb.write(output);
            output.close();
            
            /*File file1= new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            Desktop.getDesktop().open(file1);*/
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void modificar() throws IOException{
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow fila = sheet.getRow(29);

            if(fila == null){
                fila = sheet.createRow(29);
            }

            XSSFCell celda = fila.createCell(0);
            
            if(celda == null){
                celda = fila.createCell(0);
            }

            celda.setCellValue(txtFili.getText());

            file.close();

            FileOutputStream output = new FileOutputStream("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            wb.write(output);
            output.close();
            
            File file1= new File("C:\\Users\\dog_a\\Downloads\\servicio.xlsx");
            Desktop.getDesktop().open(file1);
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnBusc;
    private javax.swing.JButton btnImprimir;
    private javax.swing.JButton btnLimpi;
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel22;
    private javax.swing.JLabel jLabel23;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JSeparator jSeparator3;
    private javax.swing.JLabel lblCalle;
    private javax.swing.JLabel lblCiudad;
    private javax.swing.JLabel lblCodP;
    private javax.swing.JLabel lblColonia;
    private javax.swing.JLabel lblEstado;
    private javax.swing.JLabel lblFechaBaja;
    private javax.swing.JLabel lblFechaIng;
    private javax.swing.JLabel lblMaterno;
    private javax.swing.JLabel lblMaterno1;
    private javax.swing.JLabel lblMotivos;
    private javax.swing.JLabel lblNom;
    private javax.swing.JLabel lblNumExt;
    private javax.swing.JLabel lblPaterno;
    private javax.swing.JTable tablaPuestos;
    private javax.swing.JTextField txtCurp;
    private javax.swing.JTextField txtFili;
    private javax.swing.JTextField txtHomo;
    private javax.swing.JTextField txtRfc;
    // End of variables declaration//GEN-END:variables
}
