
package proyecto_serviciosocial;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.table.TableModel;
import proyecto_serviciosocial.panelCons;
import static proyecto_serviciosocial.panelCons.modelo;
import static proyecto_serviciosocial.panelCons.tablaPuestos;
import static proyecto_serviciosocial.panelCons.txtFili;

/**
 *
 * @author Eliud Ayala
 */
public class Agregar extends javax.swing.JFrame {

    public Agregar() {
        initComponents();
        this.setLocationRelativeTo(null);
    }
    
        conexion conect = new conexion();
        Connection con = conect.getConnection();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        txtMotivo = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        txtDel_dia = new javax.swing.JTextField();
        jLabel4 = new javax.swing.JLabel();
        txtNom = new javax.swing.JTextField();
        jLabel5 = new javax.swing.JLabel();
        txtCod = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        txtRamo = new javax.swing.JTextField();
        jLabel7 = new javax.swing.JLabel();
        txtPagad = new javax.swing.JTextField();
        jLabel8 = new javax.swing.JLabel();
        txtSueldo = new javax.swing.JTextField();
        jLabel9 = new javax.swing.JLabel();
        txtQuinq = new javax.swing.JTextField();
        jLabel10 = new javax.swing.JLabel();
        txtOtras = new javax.swing.JTextField();
        jLabel11 = new javax.swing.JLabel();
        txtTotal = new javax.swing.JTextField();
        jLabel12 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel13 = new javax.swing.JLabel();
        btnAgregar = new javax.swing.JButton();
        btnCancela = new javax.swing.JButton();
        jSeparator2 = new javax.swing.JSeparator();
        txtDel_mes = new javax.swing.JTextField();
        txtDel_ano = new javax.swing.JTextField();
        jLabel14 = new javax.swing.JLabel();
        jLabel15 = new javax.swing.JLabel();
        jLabel16 = new javax.swing.JLabel();
        jLabel17 = new javax.swing.JLabel();
        jLabel18 = new javax.swing.JLabel();
        jLabel19 = new javax.swing.JLabel();
        txtAl_dia = new javax.swing.JTextField();
        jLabel20 = new javax.swing.JLabel();
        txtAl_mes = new javax.swing.JTextField();
        txtAl_ano = new javax.swing.JTextField();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));

        jLabel1.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel1.setText("Motivo:");

        txtMotivo.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel2.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel2.setText("Del:");

        txtDel_dia.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel4.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel4.setText("Nombre:");

        txtNom.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel5.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel5.setText("Código:");

        txtCod.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel6.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel6.setText("Ramo registrado:");

        txtRamo.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel7.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel7.setText("Pagaduría registrada:");

        txtPagad.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel8.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel8.setText("Sueldo básico:");

        txtSueldo.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel9.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel9.setText("Quinquenio:");

        txtQuinq.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel10.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel10.setText("Otras percepciones:");

        txtOtras.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel11.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel11.setText("Total:");

        txtTotal.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel12.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel12.setText("Puesto");

        jSeparator1.setForeground(new java.awt.Color(0, 0, 0));

        jLabel13.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel13.setText("Agregar registro a la Tabla de bajas, reingresos, etc.");

        btnAgregar.setBackground(new java.awt.Color(153, 255, 153));
        btnAgregar.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnAgregar.setText("Agregar");
        btnAgregar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAgregarActionPerformed(evt);
            }
        });

        btnCancela.setBackground(new java.awt.Color(255, 102, 102));
        btnCancela.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnCancela.setText("Cancelar");
        btnCancela.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnCancelaActionPerformed(evt);
            }
        });

        jSeparator2.setForeground(new java.awt.Color(0, 0, 0));

        txtDel_mes.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        txtDel_ano.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel14.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel14.setText("Dia");

        jLabel15.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel15.setText("Mes");

        jLabel16.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel16.setText("Año");

        jLabel17.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel17.setText("Año");

        jLabel18.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel18.setText("Mes");

        jLabel19.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel19.setText("Dia");

        txtAl_dia.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        jLabel20.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N
        jLabel20.setText("Al:");

        txtAl_mes.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        txtAl_ano.setFont(new java.awt.Font("Arial", 0, 15)); // NOI18N

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(14, 14, 14)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel4)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(txtNom, javax.swing.GroupLayout.PREFERRED_SIZE, 227, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(64, 64, 64)
                                .addComponent(jLabel5)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtCod, javax.swing.GroupLayout.PREFERRED_SIZE, 89, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(192, 192, 192))
                            .addComponent(jLabel12)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtMotivo, javax.swing.GroupLayout.PREFERRED_SIZE, 320, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(30, 30, 30)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(txtDel_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                            .addComponent(jLabel2)
                                            .addComponent(txtDel_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtDel_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(8, 8, 8)
                                        .addComponent(jLabel14)
                                        .addGap(18, 18, 18)
                                        .addComponent(jLabel15)
                                        .addGap(18, 18, 18)
                                        .addComponent(jLabel16)))
                                .addGap(18, 18, 18)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(41, 41, 41)
                                        .addComponent(txtAl_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtAl_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(txtAl_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addGap(10, 10, 10))
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                .addComponent(jLabel19)
                                                .addGap(18, 18, 18)))
                                        .addComponent(jLabel18)
                                        .addGap(18, 18, 18)
                                        .addComponent(jLabel17))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel20)
                                        .addGap(70, 70, 70)))))
                        .addGap(0, 5, Short.MAX_VALUE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jLabel10, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGap(18, 18, 18))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(txtRamo)
                                        .addGap(38, 38, 38)))
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(txtPagad, javax.swing.GroupLayout.PREFERRED_SIZE, 102, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(txtOtras, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel6)
                                .addGap(35, 35, 35)
                                .addComponent(jLabel7))
                            .addComponent(btnAgregar, javax.swing.GroupLayout.Alignment.TRAILING))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jLabel11)
                                .addGap(3, 3, 3)
                                .addComponent(txtTotal, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(24, 24, 24)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 97, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(txtSueldo, javax.swing.GroupLayout.PREFERRED_SIZE, 110, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(24, 24, 24)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 87, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(txtQuinq, javax.swing.GroupLayout.PREFERRED_SIZE, 140, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(110, 110, 110)
                                        .addComponent(btnCancela)))
                                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 608, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(37, 37, 37)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 636, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jLabel13)
                .addGap(175, 175, 175))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel13)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel20, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtMotivo, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 15, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtDel_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtDel_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtDel_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel14, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel15, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel16, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtAl_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtAl_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtAl_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel19, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel18, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel17, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel12)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtNom, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel4)
                    .addComponent(jLabel5)
                    .addComponent(txtCod, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(31, 31, 31)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel8)
                    .addComponent(jLabel9)
                    .addComponent(jLabel7)
                    .addComponent(jLabel6))
                .addGap(6, 6, 6)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(txtPagad, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(txtQuinq, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtRamo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtSueldo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(27, 27, 27)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 15, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtOtras, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel11)
                    .addComponent(txtTotal, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 20, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnAgregar)
                    .addComponent(btnCancela, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 711, Short.MAX_VALUE)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(layout.createSequentialGroup()
                    .addGap(0, 0, Short.MAX_VALUE)
                    .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGap(0, 0, Short.MAX_VALUE)))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 381, Short.MAX_VALUE)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(layout.createSequentialGroup()
                    .addGap(0, 0, Short.MAX_VALUE)
                    .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGap(0, 0, Short.MAX_VALUE)))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btnCancelaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnCancelaActionPerformed
        this.dispose();
    }//GEN-LAST:event_btnCancelaActionPerformed

    private void btnAgregarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAgregarActionPerformed
               
        String sql = "INSERT INTO personal_sueldo (filiacion, del, al, codigo, puesto, nivel, zona, sueldo, quinquenio, otras, total, motivo)"
                + " VALUES (?,?,?,?,?,?,?,?,?,?,?,?)";
        
        /*java.sql.Date fecha1 = new java.sql.Date(0003112);
        java.sql.Date fecha2 = new java.sql.Date(19990101);*/
        
        String Del = txtDel_ano.getText() + "-" + txtDel_mes.getText() + "-" + txtDel_dia.getText();
        Del = Del.replace(" ","");
        String Al = txtAl_ano.getText() + "-" + txtAl_mes.getText() + "-" + txtAl_dia.getText();
        Al = Al.replace(" ","");
        
        try {
            System.out.println("---INICIA REGISTRO DE DATOS----");
            
            PreparedStatement ps = con.prepareStatement(sql);
            
            ps.setString(1, panelCons.txtFili.getText());
            ps.setString(2, Del);
            ps.setString(3, Al);
            ps.setString(4, "");
            ps.setString(5, "");
            ps.setString(6, "");
            ps.setString(7, "");
            ps.setDouble(8, 0);
            ps.setDouble(9, 0);
            ps.setDouble(10, 0);
            ps.setDouble(11, 0);
            ps.setString(12, txtMotivo.getText());
            
            ps.executeUpdate();
            
            JOptionPane.showMessageDialog(null, "Registro Guardado");
            llenarTabla();
            this.dispose();
            
        } catch (SQLException ex) {
            Logger.getLogger(PanelRegistro.class.getName()).log(Level.SEVERE, null, ex);
            
            System.out.println("---------------NO SE GUARDARON LOS DATOS!!!-------------");
        }
    }//GEN-LAST:event_btnAgregarActionPerformed

    public void llenarTabla(){
        Object columnas[] = new Object[12];
        
        int filas1 = tablaPuestos.getRowCount();
        for (int i = filas1 - 1; i >= 0; i--) {
            modelo.removeRow(0);
        }

        try {

            String sql = "SELECT motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras, "
                    + "sum(sueldo + quinquenio + otras) as total FROM personal_sueldo "
                    + "WHERE filiacion = '" + txtFili.getText() + "' "
                    + "AND del >= '" + sacarFechaBaja(txtFili.getText()) + "' "
                    //+ "AND motivo != '"+""+"' "
                    + "GROUP BY motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras "
                    + "ORDER BY del ASC";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                columnas[0] = rs.getString(1);
                
                Date fechaDel = rs.getDate(2);
                Date fechaAl = rs.getDate(3);
                
                columnas[1] = formatter.format(fechaDel);
                if (fechaAl == null) {
                    columnas[2] = "";
                } else {
                    columnas[2] = formatter.format(fechaAl);
                }
                columnas[3] = rs.getString(4);
                columnas[4] = rs.getString(5);
                columnas[5] = rs.getString(6);
                columnas[6] = "";
                columnas[7] = rs.getString(7);
                columnas[8] = rs.getDouble(7);
                columnas[9] = rs.getDouble(8);
                columnas[10] = rs.getDouble(9);
                columnas[11] = rs.getDouble(10);
                panelCons.modelo.addRow(columnas);
            }
            
            panelCons.tablaPuestos.setModel(panelCons.modelo);

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de agregar", "Agregar nuevo"
                    + "registro", JOptionPane.ERROR_MESSAGE);
        } 
    
    }
    
    public Date sacarFechaBaja(String filia) {
        try {
            String sql = "SELECT del FROM personal_sueldo WHERE motivo != '" + "" + "' AND "
                    + "filiacion = '" + filia + "' ORDER BY del ASC LIMIT 1";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                Date fechaBaja = rs.getDate(1);
                if (rs.getDate(1) != null) {
                    return fechaBaja;
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta de la fecha", JOptionPane.ERROR_MESSAGE);
        }
        return null;
    }
    
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Agregar.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Agregar.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Agregar.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Agregar.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Agregar().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnAgregar;
    private javax.swing.JButton btnCancela;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    public javax.swing.JTextField txtAl_ano;
    public javax.swing.JTextField txtAl_dia;
    public javax.swing.JTextField txtAl_mes;
    public javax.swing.JTextField txtCod;
    public javax.swing.JTextField txtDel_ano;
    public javax.swing.JTextField txtDel_dia;
    public javax.swing.JTextField txtDel_mes;
    public javax.swing.JTextField txtMotivo;
    public javax.swing.JTextField txtNom;
    public javax.swing.JTextField txtOtras;
    public javax.swing.JTextField txtPagad;
    public javax.swing.JTextField txtQuinq;
    public javax.swing.JTextField txtRamo;
    public javax.swing.JTextField txtSueldo;
    public javax.swing.JTextField txtTotal;
    // End of variables declaration//GEN-END:variables
}
