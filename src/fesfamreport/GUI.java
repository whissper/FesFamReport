
package fesfamreport;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.JTextField;

/**
 *
 * @author SAV2
 */
public class GUI extends JFrame
{
    private static final int WIDTH_GUI = 435;
    private static final int HEIGHT_GUI = 520;
    
    private String valSMO = "11001_";
    //
    private JLabel lblSMO;
    private JComboBox cmbbxSMO;
    
    private JLabel lblReportYear;
    private JTextField txtfReportYear;
    private JLabel lblReportMonth;
    private JTextField txtfReportMonth;
    private JLabel lblReportNumPack;
    private JTextField txtfReportNumPack;
    private JLabel lblDirPlace;
    private JTextField txtfDirPlace;
    private JLabel lblFirstRow;
    private JTextField txtfFirstRow;
    private JLabel lblLastRow;
    private JTextField txtfLastRow;
    
    private JButton btnCreateReportPackage;
    
    private JLabel lblProtocol;
    private JTextArea txtAreaProtocol;
    private JScrollPane scrollPane;
    
    public GUI()
    {
        super("FESFARM-KOMI report 2.5");
        this.makeGUI();
        
        this.cmbbxSMO.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                JComboBox box = (JComboBox)e.getSource();
                int item = (int)box.getSelectedIndex();
                if(item == 0)
                {
                    GUI.this.valSMO = "11001_";
                }
                else if(item == 1)
                {
                    GUI.this.valSMO = "11002_";
                }
            }
        });
        
        this.btnCreateReportPackage.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent me)
            {
                txtAreaProtocol.setText("");
                
                if( !txtfDirPlace.getText().equals("") &&
                    !txtfReportYear.getText().equals("") &&
                    !txtfReportMonth.getText().equals("") &&
                    !txtfReportNumPack.getText().equals("") &&
                    !txtfFirstRow.getText().equals("") &&
                    !txtfLastRow.getText().equals("")
                   )
                {
                    try
                    {
                        ReportEngine reportEng = new ReportEngine();

                        reportEng.setDirectoryPlace( txtfDirPlace.getText().replace("\\", "/") );
                        reportEng.setNamePostfix( GUI.this.valSMO,
                                                  txtfReportYear.getText().trim() + 
                                                  txtfReportMonth.getText().trim() +
                                                  txtfReportNumPack.getText().trim()
                                                );
                        reportEng.setFirstLastRows( Integer.parseInt(txtfFirstRow.getText().trim()) , 
                                                    Integer.parseInt(txtfLastRow.getText().trim()));
                        txtAreaProtocol.append( reportEng.createReport() );
                    }
                    catch(Exception e)
                    {
                        txtAreaProtocol.append("Возможно одно из полей содержит некорректные данные \nИнформация по ошибке: \n"+e);       
                    }
                }
                else
                {
                    txtAreaProtocol.append( "Пожалуйста, заполните все поля." );
                }
            }  
        
        });
    }
    
    private void makeGUI()
    {
        JPanel cp = new JPanel(null);
        this.setContentPane(cp);
        this.setSize(WIDTH_GUI, HEIGHT_GUI);
        
        this.lblSMO = new JLabel("Страховая Медицинская Организация:");
        
        String[] itemsSMO = {"СОГАЗ", "РГС"};
        this.cmbbxSMO = new JComboBox(itemsSMO);
        //this.cmbbxSMO.setEditable(true);
        
        this.lblReportYear = new JLabel("Отчетный год (последние две цифры):");
        this.txtfReportYear = new JTextField();;
        this.lblReportMonth = new JLabel("Отчетный месяц (две цифры):");;
        this.txtfReportMonth = new JTextField();;
        this.lblReportNumPack = new JLabel("Номер пакета (две цифры):");;
        this.txtfReportNumPack = new JTextField();;
        this.lblDirPlace = new JLabel("Расположение файлов:");;
        this.txtfDirPlace = new JTextField();;
        this.lblFirstRow = new JLabel("Номер первой строки с данными:");;
        this.txtfFirstRow = new JTextField();;
        this.lblLastRow = new JLabel("Номер последней строки с данными:");;
        this.txtfLastRow = new JTextField();;

        this.btnCreateReportPackage = new JButton("Сформировать отчет:");;

        this.lblProtocol = new JLabel("Протокол:");;
        this.txtAreaProtocol = new JTextArea();;
        this.scrollPane = new JScrollPane(txtAreaProtocol);
        
        JSeparator separator1 = new JSeparator();
        JSeparator separator2 = new JSeparator();
        
        cp.add(lblSMO);
        cp.add(cmbbxSMO);
        cp.add(lblReportYear);
        cp.add(txtfReportYear);
        cp.add(lblReportMonth);
        cp.add(txtfReportMonth);
        cp.add(lblReportNumPack);
        cp.add(txtfReportNumPack);
        cp.add(lblDirPlace);
        cp.add(txtfDirPlace);
        cp.add(lblFirstRow);
        cp.add(txtfFirstRow);
        cp.add(lblLastRow);
        cp.add(txtfLastRow);
        cp.add(btnCreateReportPackage);
        cp.add(lblProtocol);
        cp.add(scrollPane);
        cp.add(separator1);
        cp.add(separator2);
        
        this.lblSMO.setBounds(10, 10, 250, 25);
        this.cmbbxSMO.setBounds(250, 10, 100, 25);
        this.lblReportYear.setBounds(10, 40, 250, 25);//X, Y, Width, Height
        this.txtfReportYear.setBounds(250, 40, 100, 25);
        this.lblReportMonth.setBounds(10, 70, 250, 25);
        this.txtfReportMonth.setBounds(250, 70, 100, 25);
        this.lblReportNumPack.setBounds(10, 100, 250, 25);
        this.txtfReportNumPack.setBounds(250, 100, 100, 25);
        separator1.setBounds(10, 130, 400, 25);
        this.lblDirPlace.setBounds(10, 140, 250, 25);
        this.txtfDirPlace.setBounds(250, 140, 160, 25);
        this.lblFirstRow.setBounds(10, 170, 250, 25);
        this.txtfFirstRow.setBounds(250, 170, 100, 25);
        this.lblLastRow.setBounds(10, 200, 250, 25);
        this.txtfLastRow.setBounds(250, 200, 100, 25);
        separator2.setBounds(10, 230, 400, 25);
        this.lblProtocol.setBounds(10, 240, 100, 25);
        this.btnCreateReportPackage.setBounds(210, 240, 200, 25);
        this.scrollPane.setBounds(10, 270, 400, 200);
        
        this.setLocationRelativeTo(null);
        this.setDefaultCloseOperation(EXIT_ON_CLOSE);
        this.setVisible(true);
    }
}
