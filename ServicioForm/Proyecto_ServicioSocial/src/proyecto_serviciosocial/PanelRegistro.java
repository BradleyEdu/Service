package proyecto_serviciosocial;

import java.sql.Connection;
import java.sql.PreparedStatement;
//import java.sql.ResultSet;
import java.sql.SQLException;
//import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
//import javax.swing.table.DefaultTableModel;


public class PanelRegistro extends javax.swing.JPanel {
    
    //Conectar
    conexion conect = new conexion();
    Connection con = conect.getConnection();
    
    public PanelRegistro() {
        initComponents();
    }


    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        panelRegistro = new javax.swing.JPanel();
        lbFiliacion = new javax.swing.JLabel();
        txtFiliacion = new javax.swing.JTextField();
        lbCurp = new javax.swing.JLabel();
        txtCurp = new javax.swing.JTextField();
        lbRfc = new javax.swing.JLabel();
        txtRfc = new javax.swing.JTextField();
        lbHomo = new javax.swing.JLabel();
        txtHomo = new javax.swing.JTextField();
        lbDatos = new javax.swing.JLabel();
        lbPaterno = new javax.swing.JLabel();
        txtPaterno = new javax.swing.JTextField();
        lbMaterno = new javax.swing.JLabel();
        txtMaterno = new javax.swing.JTextField();
        lbNombre = new javax.swing.JLabel();
        txtNombre = new javax.swing.JTextField();
        lbDomicilio = new javax.swing.JLabel();
        lbCalle = new javax.swing.JLabel();
        txtCalle = new javax.swing.JTextField();
        lbColonia = new javax.swing.JLabel();
        txtColonia = new javax.swing.JTextField();
        lbNumeroExt = new javax.swing.JLabel();
        txtNumeroExt = new javax.swing.JTextField();
        lbCP = new javax.swing.JLabel();
        txtCP = new javax.swing.JTextField();
        lbCiudad = new javax.swing.JLabel();
        txtCiudad = new javax.swing.JTextField();
        lbEstado = new javax.swing.JLabel();
        txtEstado = new javax.swing.JTextField();
        lbNombre5 = new javax.swing.JLabel();
        lbIngreso = new javax.swing.JLabel();
        lbIngresoL = new javax.swing.JLabel();
        txtIngresoL = new javax.swing.JTextField();
        lbBaja = new javax.swing.JLabel();
        lbBajaL = new javax.swing.JLabel();
        txtBajaL = new javax.swing.JTextField();
        lbMotivo = new javax.swing.JLabel();
        btnGuardar = new javax.swing.JButton();
        btnCancelar = new javax.swing.JButton();
        txtDiaIng = new javax.swing.JTextField();
        txtMesIng = new javax.swing.JTextField();
        txtAnoIng = new javax.swing.JTextField();
        lbIngreso1 = new javax.swing.JLabel();
        lbIngreso2 = new javax.swing.JLabel();
        lbIngreso3 = new javax.swing.JLabel();
        lbIngreso4 = new javax.swing.JLabel();
        lbIngreso5 = new javax.swing.JLabel();
        lbIngreso6 = new javax.swing.JLabel();
        txtDiaBaja = new javax.swing.JTextField();
        txtMesBaja = new javax.swing.JTextField();
        txtAnoBaja = new javax.swing.JTextField();
        jComboBox1 = new javax.swing.JComboBox<>();
        lbMotivo1 = new javax.swing.JLabel();
        lbMotivo2 = new javax.swing.JLabel();

        panelRegistro.setBackground(new java.awt.Color(255, 255, 255));

        lbFiliacion.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbFiliacion.setForeground(new java.awt.Color(0, 0, 0));
        lbFiliacion.setText("FILIACIÓN:");

        txtFiliacion.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCurp.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCurp.setForeground(new java.awt.Color(0, 0, 0));
        lbCurp.setText("CURP:");

        txtCurp.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbRfc.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbRfc.setForeground(new java.awt.Color(0, 0, 0));
        lbRfc.setText("RFC:");

        txtRfc.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbHomo.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbHomo.setForeground(new java.awt.Color(0, 0, 0));
        lbHomo.setText("HOMOCLAVE:");

        txtHomo.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbDatos.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbDatos.setForeground(new java.awt.Color(0, 0, 0));
        lbDatos.setText("Datos Personales:");

        lbPaterno.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbPaterno.setForeground(new java.awt.Color(0, 0, 0));
        lbPaterno.setText("Apellido Paterno:");

        txtPaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbMaterno.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbMaterno.setForeground(new java.awt.Color(0, 0, 0));
        lbMaterno.setText("Apellido Materno:");

        txtMaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNombre.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNombre.setForeground(new java.awt.Color(0, 0, 0));
        lbNombre.setText("Nombre (s):");

        txtNombre.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbDomicilio.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbDomicilio.setForeground(new java.awt.Color(0, 0, 0));
        lbDomicilio.setText("Domicilio:");

        lbCalle.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCalle.setForeground(new java.awt.Color(0, 0, 0));
        lbCalle.setText("Calle:");

        txtCalle.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbColonia.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbColonia.setForeground(new java.awt.Color(0, 0, 0));
        lbColonia.setText("Colonia:");

        txtColonia.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNumeroExt.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNumeroExt.setForeground(new java.awt.Color(0, 0, 0));
        lbNumeroExt.setText("Numero Ext. e Int:");

        txtNumeroExt.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCP.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCP.setForeground(new java.awt.Color(0, 0, 0));
        lbCP.setText("CP:");

        txtCP.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCiudad.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCiudad.setForeground(new java.awt.Color(0, 0, 0));
        lbCiudad.setText("Ciudad:");

        txtCiudad.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbEstado.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbEstado.setForeground(new java.awt.Color(0, 0, 0));
        lbEstado.setText("Estado:");

        txtEstado.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNombre5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNombre5.setForeground(new java.awt.Color(0, 0, 0));
        lbNombre5.setText("PERIODO DE APORTACIONES AL FONDO DEL ISSSTE:");

        lbIngreso.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso.setText("Ingreso:");

        lbIngresoL.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngresoL.setForeground(new java.awt.Color(0, 0, 0));
        lbIngresoL.setText("Ingreso con letra:");

        txtIngresoL.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbBaja.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbBaja.setForeground(new java.awt.Color(0, 0, 0));
        lbBaja.setText("Baja:");

        lbBajaL.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbBajaL.setForeground(new java.awt.Color(0, 0, 0));
        lbBajaL.setText("Baja con letra:");

        txtBajaL.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbMotivo.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbMotivo.setForeground(new java.awt.Color(0, 0, 0));
        lbMotivo.setText("Motivo de baja:");

        btnGuardar.setBackground(new java.awt.Color(255, 255, 255));
        btnGuardar.setForeground(new java.awt.Color(255, 255, 255));
        btnGuardar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/proyecto_serviciosocial/guardar.png"))); // NOI18N
        btnGuardar.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        btnGuardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnGuardarActionPerformed(evt);
            }
        });

        btnCancelar.setBackground(new java.awt.Color(255, 255, 255));
        btnCancelar.setForeground(new java.awt.Color(255, 255, 255));
        btnCancelar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/proyecto_serviciosocial/cancel.png"))); // NOI18N
        btnCancelar.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        btnCancelar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnCancelarActionPerformed(evt);
            }
        });

        txtDiaIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtMesIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtAnoIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbIngreso1.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso1.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso1.setText("dd");

        lbIngreso2.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso2.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso2.setText("mm");

        lbIngreso3.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso3.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso3.setText("aaaa");

        lbIngreso4.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso4.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso4.setText("dd");

        lbIngreso5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso5.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso5.setText("mm");

        lbIngreso6.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso6.setForeground(new java.awt.Color(0, 0, 0));
        lbIngreso6.setText("aaaa");

        txtDiaBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtMesBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtAnoBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "RENUNCIA", "JUBILADO", "Item 3", "Item 4" }));

        lbMotivo1.setFont(new java.awt.Font("Arial", 1, 16)); // NOI18N
        lbMotivo1.setForeground(new java.awt.Color(0, 0, 0));
        lbMotivo1.setText("Guardar");

        lbMotivo2.setFont(new java.awt.Font("Arial", 1, 16)); // NOI18N
        lbMotivo2.setForeground(new java.awt.Color(0, 0, 0));
        lbMotivo2.setText("Cancelar");

        javax.swing.GroupLayout panelRegistroLayout = new javax.swing.GroupLayout(panelRegistro);
        panelRegistro.setLayout(panelRegistroLayout);
        panelRegistroLayout.setHorizontalGroup(
            panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(panelRegistroLayout.createSequentialGroup()
                .addGap(23, 23, 23)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                    .addComponent(lbFiliacion)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtFiliacion, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(lbCurp)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                    .addComponent(lbRfc)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGap(18, 18, 18)
                                    .addComponent(lbHomo)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addComponent(lbNombre5))
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                .addComponent(lbIngreso)
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGap(45, 45, 45)
                                        .addComponent(lbIngreso1)
                                        .addGap(33, 33, 33)
                                        .addComponent(lbIngreso2)
                                        .addGap(18, 18, 18)
                                        .addComponent(lbIngreso3))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                        .addGap(32, 32, 32)
                                        .addComponent(txtDiaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtMesIng, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtAnoIng, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGap(78, 78, 78)
                                .addComponent(lbIngresoL)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtIngresoL))
                            .addGroup(panelRegistroLayout.createSequentialGroup()
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(lbDatos)
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                .addComponent(lbPaterno, javax.swing.GroupLayout.Alignment.TRAILING)
                                                .addComponent(lbMaterno))
                                            .addComponent(lbNombre))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(txtPaterno, javax.swing.GroupLayout.DEFAULT_SIZE, 211, Short.MAX_VALUE)
                                            .addComponent(txtMaterno)
                                            .addComponent(txtNombre))))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 69, Short.MAX_VALUE)
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addComponent(lbCiudad)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(lbEstado)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                .addComponent(lbColonia)
                                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(lbDomicilio)
                                                    .addComponent(lbCalle)))
                                            .addComponent(lbNumeroExt))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(txtCalle)
                                            .addComponent(txtColonia)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                                .addComponent(txtNumeroExt, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 32, Short.MAX_VALUE)
                                                .addComponent(lbCP)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(txtCP, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE))))))
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                .addComponent(lbMotivo)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 187, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 402, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(panelRegistroLayout.createSequentialGroup()
                                .addComponent(lbBaja)
                                .addGap(59, 59, 59)
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addComponent(txtDiaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtMesBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(txtAnoBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGap(13, 13, 13)
                                        .addComponent(lbIngreso4)
                                        .addGap(33, 33, 33)
                                        .addComponent(lbIngreso5)
                                        .addGap(18, 18, 18)
                                        .addComponent(lbIngreso6)))
                                .addGap(108, 108, 108)
                                .addComponent(lbBajaL)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtBajaL))
                            .addGroup(panelRegistroLayout.createSequentialGroup()
                                .addGap(0, 0, Short.MAX_VALUE)
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                        .addComponent(btnGuardar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(56, 56, 56))
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                        .addComponent(lbMotivo1)
                                        .addGap(68, 68, 68)))
                                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGap(74, 74, 74)
                                        .addComponent(btnCancelar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(panelRegistroLayout.createSequentialGroup()
                                        .addGap(88, 88, 88)
                                        .addComponent(lbMotivo2)))
                                .addGap(297, 297, 297)))
                        .addGap(41, 41, 41))))
        );
        panelRegistroLayout.setVerticalGroup(
            panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(panelRegistroLayout.createSequentialGroup()
                .addGap(23, 23, 23)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbFiliacion)
                    .addComponent(txtFiliacion, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lbCurp)
                    .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(37, 37, 37)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbRfc)
                    .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lbHomo)
                    .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(41, 41, 41)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbDatos)
                    .addComponent(lbDomicilio))
                .addGap(18, 18, 18)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbPaterno)
                            .addComponent(txtPaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbMaterno)
                            .addComponent(txtMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbNombre)
                            .addComponent(txtNombre, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbCalle)
                            .addComponent(txtCalle, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbColonia)
                            .addComponent(txtColonia, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbNumeroExt)
                            .addComponent(txtNumeroExt, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(lbCP)
                            .addComponent(txtCP, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbEstado)
                    .addComponent(txtEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(lbCiudad)
                    .addComponent(txtCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(47, 47, 47)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addComponent(lbNombre5)
                        .addGap(46, 46, 46)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbIngreso)
                            .addComponent(lbIngresoL)
                            .addComponent(txtIngresoL, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbIngreso1)
                            .addComponent(lbIngreso2)
                            .addComponent(lbIngreso3))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtDiaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtMesIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtAnoIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(18, 18, 18)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbBaja)
                        .addComponent(lbBajaL)
                        .addComponent(txtBajaL, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(panelRegistroLayout.createSequentialGroup()
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(lbIngreso4)
                            .addComponent(lbIngreso5)
                            .addComponent(lbIngreso6))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtDiaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtMesBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtAnoBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(48, 48, 48)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbMotivo)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(55, 55, 55)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btnCancelar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnGuardar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(lbMotivo1)
                    .addComponent(lbMotivo2))
                .addContainerGap(19, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(panelRegistro, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(panelRegistro, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );
    }// </editor-fold>//GEN-END:initComponents

    private void btnGuardarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnGuardarActionPerformed
        String sql, filiacion,curp,rfc,homoclave,apellido_paterno, apellido_materno,nombre,calle,numero,colonia,
                codigo_postal,ciudad,estado,ingreso,ingreso_letra,baja,baja_letra,motivo_baja;
        String fechaIng,fechaBaja;
        filiacion = txtFiliacion.getText();
        curp = txtCurp.getText();
        rfc = txtRfc.getText();
        homoclave = txtHomo.getText();
        apellido_paterno = txtPaterno.getText();
        apellido_materno = txtMaterno.getText();
        nombre = txtNombre.getText();
        calle = txtCalle.getText();
        numero = txtNumeroExt.getText();
        colonia = txtColonia.getText();
        codigo_postal = txtCP.getText();
        ciudad = txtCiudad.getText();
        estado = txtEstado.getText();
        fechaIng=txtDiaIng.getText()+"/"+txtMesIng.getText()+"/"+txtAnoIng.getText();
        ingreso = fechaIng;
        ingreso_letra = txtNombre.getText();
        fechaBaja=txtDiaBaja.getText()+"/"+txtMesBaja.getText()+"/"+txtAnoBaja.getText();
        baja = fechaBaja;
        baja_letra = txtNombre.getText();
        motivo_baja = txtNombre.getText();
        
        //insertar en la tabla PERSONAL (campos de la tabla)valores (variables)
        sql = "INSERT INTO personal (filiacion) VALUES (?)";
        try {
            System.out.println("---INICIA REGISTRO DE DATOS----");
            PreparedStatement ps = con.prepareStatement(sql);
            ps.setString(1, filiacion);
//            ps.setString(2, curp);
//            ps.setString(3, rfc);
            ps.executeUpdate();
            //verTabla();
            JOptionPane.showMessageDialog(null, "Registro Guardado");
            //limpiar();
        } catch (SQLException ex) {
            Logger.getLogger(PanelRegistro.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("NO SE GUARDARON LOS DATOS!!!");
        }
        

    }//GEN-LAST:event_btnGuardarActionPerformed

    private void btnCancelarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnCancelarActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_btnCancelarActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnCancelar;
    private javax.swing.JButton btnGuardar;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JLabel lbBaja;
    private javax.swing.JLabel lbBajaL;
    private javax.swing.JLabel lbCP;
    private javax.swing.JLabel lbCalle;
    private javax.swing.JLabel lbCiudad;
    private javax.swing.JLabel lbColonia;
    private javax.swing.JLabel lbCurp;
    private javax.swing.JLabel lbDatos;
    private javax.swing.JLabel lbDomicilio;
    private javax.swing.JLabel lbEstado;
    private javax.swing.JLabel lbFiliacion;
    private javax.swing.JLabel lbHomo;
    private javax.swing.JLabel lbIngreso;
    private javax.swing.JLabel lbIngreso1;
    private javax.swing.JLabel lbIngreso2;
    private javax.swing.JLabel lbIngreso3;
    private javax.swing.JLabel lbIngreso4;
    private javax.swing.JLabel lbIngreso5;
    private javax.swing.JLabel lbIngreso6;
    private javax.swing.JLabel lbIngresoL;
    private javax.swing.JLabel lbMaterno;
    private javax.swing.JLabel lbMotivo;
    private javax.swing.JLabel lbMotivo1;
    private javax.swing.JLabel lbMotivo2;
    private javax.swing.JLabel lbNombre;
    private javax.swing.JLabel lbNombre5;
    private javax.swing.JLabel lbNumeroExt;
    private javax.swing.JLabel lbPaterno;
    private javax.swing.JLabel lbRfc;
    private javax.swing.JPanel panelRegistro;
    private javax.swing.JTextField txtAnoBaja;
    private javax.swing.JTextField txtAnoIng;
    private javax.swing.JTextField txtBajaL;
    private javax.swing.JTextField txtCP;
    private javax.swing.JTextField txtCalle;
    private javax.swing.JTextField txtCiudad;
    private javax.swing.JTextField txtColonia;
    private javax.swing.JTextField txtCurp;
    private javax.swing.JTextField txtDiaBaja;
    private javax.swing.JTextField txtDiaIng;
    private javax.swing.JTextField txtEstado;
    private javax.swing.JTextField txtFiliacion;
    private javax.swing.JTextField txtHomo;
    private javax.swing.JTextField txtIngresoL;
    private javax.swing.JTextField txtMaterno;
    private javax.swing.JTextField txtMesBaja;
    private javax.swing.JTextField txtMesIng;
    private javax.swing.JTextField txtNombre;
    private javax.swing.JTextField txtNumeroExt;
    private javax.swing.JTextField txtPaterno;
    private javax.swing.JTextField txtRfc;
    // End of variables declaration//GEN-END:variables
}
