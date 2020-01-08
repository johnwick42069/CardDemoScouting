package com.company;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.Font;
import java.awt.event.*;
import javax.swing.*;
import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JComboBox;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Container;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import java.io.*;
import java.lang.management.ManagementFactory;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class Main<MNS> implements ItemListener {

    //RestartFunction Info
    public static final String SUN_JAVA_COMMAND = "sun.java.command";

    JPanel cards; //a panel that uses CardLayout
    //Grid Layout Settings
    static final String gapList[] = {"0", "10", "15", "20"};
    final static int maxGap = 20;
    JComboBox horGapComboBox;
    JComboBox verGapComboBox;
    JButton applyButton = new JButton("Apply gaps");
    GridLayout experimentLayout = new GridLayout(0, 2);

    //forward/backward
    JButton EXIT = new JButton("X");
    JButton NEXTpregame = new JButton("START");
    JButton NEXTpregame2 = new JButton("START");
    JButton PREVsandstorm = new JButton("BACK");
    JButton NEXTsandstorm = new JButton("NEXT");
    JButton PREVteleop = new JButton("BACK");
    JButton NEXTteleop = new JButton("NEXT");
    JButton PREVendgame = new JButton("<");
    JButton SUBMIT = new JButton("+");
    //TIMER
    boolean first = true;
    int timer = 0;
    int t = 0;

    //PHASE 1
    JSpinner AMNS = new JSpinner(); //Match Number
    String SAMNS = new String("0");
    JSpinner ATNS = new JSpinner(); //Team Number
    String SATNS = new String("0");
    //PHASE 2
    JSpinner BRC3 = new JSpinner(); //Sandstorm Rocket 3 Cargo
    String SBRC3 = new String("0");
    JSpinner BRC2 = new JSpinner(); //Sandstorm Rocket 2 Cargo
    String SBRC2 = new String("0");
    JSpinner BRC1 = new JSpinner(); //Sandstorm Rocket 1 Cargo
    String SBRC1 = new String("0");
    JSpinner BRCS = new JSpinner(); //Ship Cargo
    String SBRCS = new String("0");
    JSpinner BRH3 = new JSpinner(); //Sandstorm Rocket 3 Hatch
    String SBRH3 = new String("0");
    JSpinner BRH2 = new JSpinner(); //Sandstorm Rocket 2 Hatch
    String SBRH2 = new String("0");
    JSpinner BRH1 = new JSpinner(); //Sandstorm Rocket 1 Hatch
    String SBRH1 = new String("0");
    JSpinner BRHS = new JSpinner(); //Ship Hatch
    String SBRHS = new String("0");
    JRadioButton BGB = new JRadioButton("Ground?"); //Sandstorm Button to qualify ground pick-up
    String SBGB = new String("false");
    JRadioButton BCB = new JRadioButton("Cargo"); //Sandstorm Ground Pickup Cargo
    String SBCB = new String("false");
    JRadioButton BHB = new JRadioButton("Hatch"); //Sandstorm Ground Pickup Hatch
    String SBHB = new String("false");
    //PHASE 3
    JSpinner CRC3 = new JSpinner(); //Tele-op Rocket 3 Cargo
    String SCRC3 = new String("0");
    JSpinner CRC2 = new JSpinner(); //Tele-op Rocket 2 Cargo
    String SCRC2 = new String("0");
    JSpinner CRC1 = new JSpinner(); //Tele-op Rocket 1 Cargo
    String SCRC1 = new String("0");
    JSpinner CRCS = new JSpinner(); //Tele-op Ship Cargo
    String SCRCS = new String("0");
    JSpinner CRH3 = new JSpinner(); //Tele-op Rocket 3 Hatch
    String SCRH3 = new String("0");
    JSpinner CRH2 = new JSpinner(); //Tele-op Rocket 2 Hatch
    String SCRH2 = new String("0");
    JSpinner CRH1 = new JSpinner(); //Tele-op Rocket 1 Hatch
    String SCRH1 = new String("0");
    JSpinner CRHS = new JSpinner(); //Tele-op Ship Hatch
    String SCRHS = new String("0");
    JRadioButton CGB = new JRadioButton("Ground?"); //Tele-op Button to qualify ground pick-up
    String SCGB = new String("false");
    JRadioButton CCB = new JRadioButton("Cargo"); //Tele-op Ground Pickup Cargo
    String SCCB = new String("false");
    JRadioButton CHB = new JRadioButton("Hatch"); //Tele-op Ground Pickup Hatch
    String SCHB = new String("false");
    //PHASE 4
    JRadioButton DL3 = new JRadioButton("lvl 3"); //Climb Level 3
    String SDL3 = new String("false");
    JRadioButton DL2 = new JRadioButton("lvl 2"); //Climb Level 2
    String SDL2 = new String("false");
    JRadioButton DL1 = new JRadioButton("lvl 1"); //Climb Level 1
    String SDL1 = new String("false");
    JRadioButton DFB = new JRadioButton("Fouls?"); //Fouls
    String SDFB = new String("false");
    JTextField DDT = new JTextField(); //Defense Write
    String SDDT = new String("---");
    JTextField DNT = new JTextField(); //Notes Write
    String SDNT = new String("---");

    //DesktopPath
    String UserName = System.getProperty("user.name");
    //ICONS
    ImageIcon homescreen = new ImageIcon("C:" + '\\' + "Users" + '\\' + UserName + '\\' + "Desktop" + '\\' + "Scout5442-19" + '\\' + "homescreen.jpg");
    public void initGaps() {
        horGapComboBox = new JComboBox(gapList);
        verGapComboBox = new JComboBox(gapList);
    }

    public void addComponentToPane(Container pane) {
        //TEST 3/15/2019
        System.out.println(UserName);
        File jarDir = new File(ClassLoader.getSystemClassLoader().getResource(".").getPath());
        System.out.println(jarDir.getAbsolutePath());
        //INITIAL BUTTON DISABLE
        BCB.setEnabled(false);
        BHB.setEnabled(false);
        CCB.setEnabled(false);
        CHB.setEnabled(false);

        //Put the JComboBox in a JPanel to get a nicer look.
        JPanel comboBoxPane = new JPanel(); //use FlowLayout
        //Create the "cards".

        //INITIALIZE
        JPanel controls = new JPanel();
        JButton b = new JButton("Just fake button");
        final Dimension buttonSize = b.getPreferredSize();
        //GridBag (Delete Later?)
        GridBagConstraints c = new GridBagConstraints();
        c.fill = GridBagConstraints.CENTER;
        c.ipadx = 0;
        c.ipady = 0;
        c.weightx = .5;
        c.weighty = .5;
        //
        int x = 0;
        int y = 0;
        //Style
        final Font buttonFont = new Font(Font.SANS_SERIF, Font.BOLD, 18);
        Font titleFont = new Font(Font.SANS_SERIF, Font.BOLD, 30);
        Font TextFont = new Font(Font.SANS_SERIF, Font.PLAIN, 16);


        //PREGAME
        final JPanel pregame = new JPanel();
        //big picture
        JLabel bigPicture = new JLabel(homescreen);
        bigPicture.setBounds(0, 0, 350, 320);
        pregame.add(bigPicture);
        pregame.setBackground(Color.BLACK);
        Insets insets = pregame.getInsets();
        pregame.setLayout(null);
        pregame.setPreferredSize(new Dimension((int) (buttonSize.getWidth() * 5), (int) (buttonSize.getHeight() * 12)));
        //Picture[insert]
        //MATCH LABEL
        x = (int) (buttonSize.getWidth() * 3);
        y = (int) (buttonSize.getHeight() * 1.8);
        JLabel PregameMatchNumber = new JLabel("MATCH #", SwingConstants.RIGHT);
        PregameMatchNumber.setBounds(x, y, 100, 50);
        PregameMatchNumber.setForeground(Color.ORANGE);
        PregameMatchNumber.setFont(buttonFont);
        pregame.add(PregameMatchNumber);
        //TEAM LABEL
        x = (int) (buttonSize.getWidth() * 3);
        y = (int) (buttonSize.getHeight() * 4.3);
        JLabel PregameTeamNumber = new JLabel("TEAM #", SwingConstants.RIGHT);
        PregameTeamNumber.setBounds(x, y, 100, 50);
        PregameTeamNumber.setForeground(Color.ORANGE);
        PregameTeamNumber.setFont(buttonFont);
        pregame.add(PregameTeamNumber);
        //MATCH SPINNER
        x = (int) (buttonSize.getWidth() * 4);
        y = (int) (buttonSize.getHeight() * 2);
        AMNS.setBounds(x, y, 100, 40);
        AMNS.setForeground(Color.BLACK);
        pregame.add(AMNS);
        //TEAM SPINNER
        x = (int) (buttonSize.getWidth() * 4);
        y = (int) (buttonSize.getHeight() * 4.5);
        ATNS.setBounds(x, y, 100, 40);
        pregame.add(ATNS);
        x = (int) (buttonSize.getWidth() * 4);
        y = (int) (buttonSize.getHeight() * 9);
        if (first = true) {
            NEXTpregame.setBackground(Color.ORANGE);
            NEXTpregame.setForeground(Color.BLACK);
            NEXTpregame.setFont(buttonFont);
            NEXTpregame.setBounds(x, y, 100, 40);
            pregame.add(NEXTpregame);
        } else if (first = false) {
            NEXTpregame2.setBackground(Color.ORANGE);
            NEXTpregame2.setForeground(Color.BLACK);
            NEXTpregame2.setFont(buttonFont);
            NEXTpregame2.setBounds(x, y, 100, 40);
            pregame.add(NEXTpregame2);
        }

        //SANDSTORM
        final JPanel sandstorm = new JPanel();
        sandstorm.setBackground(Color.BLACK);
        sandstorm.setLayout(null);
        sandstorm.setPreferredSize(new Dimension((int) (buttonSize.getWidth() * 5) + maxGap, (int) (buttonSize.getHeight() * 7) + maxGap * 2));
        //GreyPanel
        JPanel SandstormGrey = new JPanel();
        SandstormGrey.setLayout(new GridLayout(5, 1, 0, 0));
        SandstormGrey.setBackground(Color.GRAY);
        x = (int) (buttonSize.getWidth() * 3.51);
        y = (int) (buttonSize.getHeight() * 0);
        SandstormGrey.setBounds(x, y, 200, 273);
        sandstorm.add(SandstormGrey);
        //Sandstorm Label
        JLabel sandstormTitle = new JLabel("SANDSTORM");
        x = (int) (buttonSize.getWidth() * 1.275);
        y = (int) (buttonSize.getHeight() * 0);
        sandstormTitle.setFont(titleFont);
        sandstormTitle.setForeground(Color.ORANGE);
        sandstormTitle.setBounds(x, y, 200, 50);
        sandstorm.add(sandstormTitle);
        //Rocket3 Label
        JLabel sandstormRocket3 = new JLabel(" ROCKET 3");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 3.5);
        sandstormRocket3.setFont(buttonFont);
        sandstormRocket3.setForeground(Color.ORANGE);
        sandstormRocket3.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormRocket3);
        //Rocket2 Label
        JLabel sandstormRocket2 = new JLabel(" ROCKET 2");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 5.5);
        sandstormRocket2.setFont(buttonFont);
        sandstormRocket2.setForeground(Color.ORANGE);
        sandstormRocket2.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormRocket2);
        //Rocket1 Label
        JLabel sandstormRocket1 = new JLabel(" ROCKET 1");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 7.5);
        sandstormRocket1.setFont(buttonFont);
        sandstormRocket1.setForeground(Color.ORANGE);
        sandstormRocket1.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormRocket1);
        //CargoShip Label
        JLabel sandstormCargoShip = new JLabel("           SHIP");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 9.5);
        sandstormCargoShip.setFont(buttonFont);
        sandstormCargoShip.setForeground(Color.ORANGE);
        sandstormCargoShip.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormCargoShip);
        //Cargo Label
        JLabel sandstormCargo = new JLabel("Cargo");
        x = (int) (buttonSize.getWidth() * 1.3);
        y = (int) (buttonSize.getHeight() * 1.75);
        sandstormCargo.setFont(buttonFont);
        sandstormCargo.setForeground(Color.ORANGE);
        sandstormCargo.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormCargo);
        //Hatch Label
        JLabel sandstormHatch = new JLabel("Hatch");
        x = (int) (buttonSize.getWidth() * 2.4);
        y = (int) (buttonSize.getHeight() * 1.75);
        sandstormHatch.setFont(buttonFont);
        sandstormHatch.setForeground(Color.ORANGE);
        sandstormHatch.setBounds(x, y, 100, 50);
        sandstorm.add(sandstormHatch);
        //------------------------
        //        ROCKETS
        //------------------------
        //Spinner Cargo Rocket 3
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 3.5);
        BRC3.setFont(buttonFont);
        BRC3.setForeground(Color.ORANGE);
        BRC3.setBounds(x, y, 100, 40);
        sandstorm.add(BRC3);
        //Spinner Cargo Rocket 2
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 5.5);
        BRC2.setFont(buttonFont);
        BRC2.setForeground(Color.ORANGE);
        BRC2.setBounds(x, y, 100, 40);
        sandstorm.add(BRC2);
        //Spinner Cargo Rocket 1
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 7.5);
        BRC1.setFont(buttonFont);
        BRC1.setForeground(Color.ORANGE);
        BRC1.setBounds(x, y, 100, 40);
        sandstorm.add(BRC1);
        //Spinner Cargo Ship
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 9.5);
        BRCS.setFont(buttonFont);
        BRCS.setForeground(Color.ORANGE);
        BRCS.setBounds(x, y, 100, 40);
        sandstorm.add(BRCS);
        //Spinner Hatch Rocket 3
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 3.5);
        BRH3.setFont(buttonFont);
        BRH3.setForeground(Color.ORANGE);
        BRH3.setBounds(x, y, 100, 40);
        sandstorm.add(BRH3);
        //Spinner Hatch Rocket 2
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 5.5);
        BRH2.setFont(buttonFont);
        BRH2.setForeground(Color.ORANGE);
        BRH2.setBounds(x, y, 100, 40);
        sandstorm.add(BRH2);
        //Spinner Hatch Rocket 1
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 7.5);
        BRH1.setFont(buttonFont);
        BRH1.setForeground(Color.ORANGE);
        BRH1.setBounds(x, y, 100, 40);
        sandstorm.add(BRH1);
        //Spinner Hatch Ship
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 9.5);
        BRHS.setFont(buttonFont);
        BRHS.setForeground(Color.ORANGE);
        BRHS.setBounds(x, y, 100, 40);
        sandstorm.add(BRHS);
        //------
        //SANDSTORMLAYOUT PAINT
        //------
        //JLabel Filler
        final JLabel SandstormTimer = new JLabel("NULL", SwingConstants.CENTER);
        SandstormTimer.setFont(titleFont);
        SandstormGrey.setBackground(Color.WHITE);
        sandstorm.add(SandstormTimer);
        SandstormGrey.add(SandstormTimer);
        //BGB
        BGB.setBackground(Color.GRAY);
        BGB.setForeground(Color.BLACK);
        BGB.setFont(buttonFont);
        sandstorm.add(BGB);
        SandstormGrey.add(BGB);
        //BCB
        BCB.setBackground(Color.DARK_GRAY);
        BCB.setForeground(Color.WHITE);
        BCB.setFont(buttonFont);
        sandstorm.add(BCB);
        SandstormGrey.add(BCB);
        //BHB
        BHB.setBackground(Color.DARK_GRAY);
        BHB.setForeground(Color.WHITE);
        BHB.setFont(buttonFont);
        sandstorm.add(BHB);
        SandstormGrey.add(BHB);
        JLabel SandstormFiller = new JLabel();
        SandstormGrey.setBackground(Color.GRAY);
        sandstorm.add(SandstormFiller);
        SandstormGrey.add(SandstormFiller);
        //------
        //SANDSTORMLAYOUT END
        //------
        //PREVsandstorm
        x = (int) (buttonSize.getWidth() * 3.51);
        y = (int) (buttonSize.getHeight() * 10.5);
        PREVsandstorm.setBackground(Color.LIGHT_GRAY);
        PREVteleop.setForeground(Color.BLACK);
        PREVsandstorm.setFont(buttonFont);
        PREVsandstorm.setBounds(x, y, 100, 40);
        sandstorm.add(PREVsandstorm);
        //NEXTsandstorm
        x = (int) (buttonSize.getWidth() * 4.34);
        y = (int) (buttonSize.getHeight() * 10.5);
        NEXTsandstorm.setBackground(Color.LIGHT_GRAY);
        NEXTsandstorm.setForeground(Color.DARK_GRAY);
        NEXTsandstorm.setFont(buttonFont);
        NEXTsandstorm.setBounds(x, y, 100, 40);
        sandstorm.add(NEXTsandstorm);


        //TELEOP
        JPanel teleop = new JPanel();
        teleop.setBackground(Color.BLACK);
        teleop.setLayout(null);
        teleop.setPreferredSize(new Dimension((int) (buttonSize.getWidth() * 5) + maxGap, (int) (buttonSize.getHeight() * 7) + maxGap * 2));
        //GreyPanel
        JPanel TeleopGrey = new JPanel();
        TeleopGrey.setLayout(new GridLayout(5, 1, 0, 0));
        TeleopGrey.setBackground(Color.GRAY);
        x = (int) (buttonSize.getWidth() * 3.51);
        y = (int) (buttonSize.getHeight() * 0);
        TeleopGrey.setBounds(x, y, 200, 273);
        teleop.add(TeleopGrey);
        //Sandstorm Label
        JLabel teleopTitle = new JLabel("TELEOP");
        x = (int) (buttonSize.getWidth() * 1.625);
        y = (int) (buttonSize.getHeight() * 0);
        teleopTitle.setFont(titleFont);
        teleopTitle.setForeground(Color.ORANGE);
        teleopTitle.setBounds(x, y, 200, 50);
        teleop.add(teleopTitle);
        //Rocket3 Label
        JLabel teleopRocket3 = new JLabel(" ROCKET 3");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 3.5);
        teleopRocket3.setFont(buttonFont);
        teleopRocket3.setForeground(Color.ORANGE);
        teleopRocket3.setBounds(x, y, 100, 50);
        teleop.add(teleopRocket3);
        //Rocket2 Label
        JLabel teleopRocket2 = new JLabel(" ROCKET 2");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 5.5);
        teleopRocket2.setFont(buttonFont);
        teleopRocket2.setForeground(Color.ORANGE);
        teleopRocket2.setBounds(x, y, 100, 50);
        teleop.add(teleopRocket2);
        //Rocket1 Label
        JLabel teleopRocket1 = new JLabel(" ROCKET 1");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 7.5);
        teleopRocket1.setFont(buttonFont);
        teleopRocket1.setForeground(Color.ORANGE);
        teleopRocket1.setBounds(x, y, 100, 50);
        teleop.add(teleopRocket1);
        //Rocket1 Label
        JLabel teleopCargoShip = new JLabel("           SHIP");
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 9.5);
        teleopCargoShip.setFont(buttonFont);
        teleopCargoShip.setForeground(Color.ORANGE);
        teleopCargoShip.setBounds(x, y, 100, 50);
        teleop.add(teleopCargoShip);
        //Cargo Label
        JLabel teleopCargo = new JLabel("Cargo");
        x = (int) (buttonSize.getWidth() * 1.3);
        y = (int) (buttonSize.getHeight() * 1.75);
        teleopCargo.setFont(buttonFont);
        teleopCargo.setForeground(Color.ORANGE);
        teleopCargo.setBounds(x, y, 100, 50);
        teleop.add(teleopCargo);
        //Hatch Label
        JLabel teleopHatch = new JLabel("Hatch");
        x = (int) (buttonSize.getWidth() * 2.4);
        y = (int) (buttonSize.getHeight() * 1.75);
        teleopHatch.setFont(buttonFont);
        teleopHatch.setForeground(Color.ORANGE);
        teleopHatch.setBounds(x, y, 100, 50);
        teleop.add(teleopHatch);
        //------------------------
        //        ROCKETS
        //------------------------
        //Spinner Cargo Rocket 3
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 3.5);
        CRC3.setFont(buttonFont);
        CRC3.setForeground(Color.ORANGE);
        CRC3.setBounds(x, y, 100, 40);
        teleop.add(CRC3);
        //Spinner Cargo Rocket 2
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 5.5);
        CRC2.setFont(buttonFont);
        CRC2.setForeground(Color.ORANGE);
        CRC2.setBounds(x, y, 100, 40);
        teleop.add(CRC2);
        //Spinner Cargo Rocket 1
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 7.5);
        CRC1.setFont(buttonFont);
        CRC1.setForeground(Color.ORANGE);
        CRC1.setBounds(x, y, 100, 40);
        teleop.add(CRC1);
        //Spinner Cargo Ship
        x = (int) (buttonSize.getWidth() * 1.15);
        y = (int) (buttonSize.getHeight() * 9.5);
        CRCS.setFont(buttonFont);
        CRCS.setForeground(Color.ORANGE);
        CRCS.setBounds(x, y, 100, 40);
        teleop.add(CRCS);
        //Spinner Hatch Rocket 3
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 3.5);
        CRH3.setFont(buttonFont);
        CRH3.setForeground(Color.ORANGE);
        CRH3.setBounds(x, y, 100, 40);
        teleop.add(CRH3);
        //Spinner Hatch Rocket 2
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 5.5);
        CRH2.setFont(buttonFont);
        CRH2.setForeground(Color.ORANGE);
        CRH2.setBounds(x, y, 100, 40);
        teleop.add(CRH2);
        //Spinner Hatch Rocket 1
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 7.5);
        CRH1.setFont(buttonFont);
        CRH1.setForeground(Color.ORANGE);
        CRH1.setBounds(x, y, 100, 40);
        teleop.add(CRH1);
        //Spinner Hatch Ship
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 9.5);
        CRHS.setFont(buttonFont);
        CRHS.setForeground(Color.ORANGE);
        CRHS.setBounds(x, y, 100, 40);
        teleop.add(CRHS);
        //------
        //SANDSTORMLAYOUT PAINT
        //------
        //JLabel Filler
        JLabel TeleopFiller1 = new JLabel();
        teleop.add(TeleopFiller1);
        TeleopGrey.add(TeleopFiller1);
        //BGB
        CGB.setBackground(Color.GRAY);
        CGB.setForeground(Color.BLACK);
        CGB.setFont(buttonFont);
        teleop.add(CGB);
        TeleopGrey.add(CGB);
        //BCB
        CCB.setBackground(Color.DARK_GRAY);
        CCB.setForeground(Color.WHITE);
        CCB.setFont(buttonFont);
        teleop.add(CCB);
        TeleopGrey.add(CCB);
        //BHB
        CHB.setBackground(Color.DARK_GRAY);
        CHB.setForeground(Color.WHITE);
        CHB.setFont(buttonFont);
        teleop.add(CHB);
        TeleopGrey.add(CHB);
        //JLabelFiller
        JLabel TeleopFiller2 = new JLabel();
        teleop.add(TeleopFiller2);
        TeleopGrey.add(TeleopFiller2);
        //------
        //TELEOPLAYOUT END
        //------
        //PREVteleop
        x = (int) (buttonSize.getWidth() * 3.51);
        y = (int) (buttonSize.getHeight() * 10.5);
        PREVteleop.setBackground(Color.LIGHT_GRAY);
        PREVteleop.setForeground(Color.BLACK);
        PREVteleop.setFont(buttonFont);
        PREVteleop.setBounds(x, y, 100, 40);
        teleop.add(PREVteleop);
        //NEXTteleop
        x = (int) (buttonSize.getWidth() * 4.34);
        y = (int) (buttonSize.getHeight() * 10.5);
        NEXTteleop.setBackground(Color.ORANGE);
        NEXTteleop.setForeground(Color.BLACK);
        NEXTteleop.setFont(buttonFont);
        NEXTteleop.setBounds(x, y, 100, 40);
        teleop.add(NEXTteleop);

        //ENDGAME
        JPanel endgame = new JPanel();
        endgame.setBackground(Color.BLACK);
        endgame.setLayout(null);
        endgame.setPreferredSize(new Dimension((int) (buttonSize.getWidth() * 5) + maxGap, (int) (buttonSize.getHeight() * 7) + maxGap * 2));
        JPanel EndgameGrey = new JPanel();
        EndgameGrey.setLayout(new GridLayout(7, 1, 0, 0));
        EndgameGrey.setBackground(Color.GRAY);
        x = (int) (buttonSize.getWidth() * 0);
        y = (int) (buttonSize.getHeight() * 0);
        EndgameGrey.setBounds(x, y, 200, 315);
        endgame.add(EndgameGrey);
        //GREYPANEL START
        //JLabel Filler
        JLabel EndgameFiller1 = new JLabel();
        endgame.add(EndgameFiller1);
        EndgameGrey.add(EndgameFiller1);
        JLabel EndgameFiller2 = new JLabel("CLIMB?");
        EndgameFiller2.setFont(titleFont);
        endgame.add(EndgameFiller2);
        EndgameGrey.add(EndgameFiller2);
        //DL3
        DL3.setBackground(Color.DARK_GRAY);
        DL3.setForeground(Color.WHITE);
        DL3.setFont(buttonFont);
        endgame.add(DL3);
        EndgameGrey.add(DL3);
        //DL2
        DL2.setBackground(Color.DARK_GRAY);
        DL2.setForeground(Color.WHITE);
        DL2.setFont(buttonFont);
        endgame.add(DL2);
        EndgameGrey.add(DL2);
        //DL1
        DL1.setBackground(Color.DARK_GRAY);
        DL1.setForeground(Color.WHITE);
        DL1.setFont(buttonFont);
        endgame.add(DL1);
        EndgameGrey.add(DL1);
        //DFB
        DFB.setBackground(Color.GRAY);
        DFB.setForeground(Color.BLACK);
        DFB.setFont(buttonFont);
        endgame.add(DFB);
        EndgameGrey.add(DFB);
        //JLabelFiller
        JLabel EndgameFiller3 = new JLabel();
        endgame.add(EndgameFiller3);
        EndgameGrey.add(EndgameFiller3);
        //GREYPANEL END
        //DDT LABEL
        JLabel EndgameDDTLabel = new JLabel("DEFENSE? (include fouls)");
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * .5);
        EndgameDDTLabel.setFont(buttonFont);
        EndgameDDTLabel.setForeground(Color.ORANGE);
        EndgameDDTLabel.setBounds(x, y, 300, 50);
        endgame.add(EndgameDDTLabel);
        //DDT
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 2);
        DDT.setBounds(x, y, 300, 75);
        DDT.setFont(TextFont);
        endgame.add(DDT);
        //DNT LABEL
        JLabel EndgameDNTLabel = new JLabel("OTHER NOTES?");
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 5);
        EndgameDNTLabel.setFont(buttonFont);
        EndgameDNTLabel.setForeground(Color.ORANGE);
        EndgameDNTLabel.setBounds(x, y, 300, 50);
        endgame.add(EndgameDNTLabel);
        //DNT
        x = (int) (buttonSize.getWidth() * 2.25);
        y = (int) (buttonSize.getHeight() * 6.5);
        DNT.setBounds(x, y, 300, 75);
        DNT.setFont(TextFont);
        endgame.add(DNT);
        ///////////////////////////////////////
        x = (int) (buttonSize.getWidth() * 3.51);
        y = (int) (buttonSize.getHeight() * 10.5);
        PREVendgame.setBackground(Color.LIGHT_GRAY);
        PREVendgame.setForeground(Color.BLACK);
        PREVendgame.setFont(buttonFont);
        PREVendgame.setBounds(x, y, 100, 40);
        endgame.add(PREVendgame);
        //NEXTteleop
        x = (int) (buttonSize.getWidth() * 4.34);
        y = (int) (buttonSize.getHeight() * 10.5);
        SUBMIT.setBackground(Color.ORANGE);
        SUBMIT.setForeground(Color.BLACK);
        SUBMIT.setFont(buttonFont);
        SUBMIT.setBounds(x, y, 100, 40);
        endgame.add(SUBMIT);

        //ACTION LISTENERS
        //---------------------------------------
        //              STAGE ONE
        //---------------------------------------
        AMNS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SAMNS = (AMNS.getValue() + "");
            }
        });
        ATNS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SATNS = (ATNS.getValue() + "");
            }
        });
        //---------------------------------------
        //              STAGE TWO
        //---------------------------------------
        BRC3.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRC3 = (BRC3.getValue() + "");
            }
        });
        BRC2.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRC2 = (BRC2.getValue() + "");
            }
        });
        BRC1.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRC1 = (BRC1.getValue() + "");
            }
        });
        BRCS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRCS = (BRCS.getValue() + "");
            }
        });
        BRH3.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRH3 = (BRH3.getValue() + "");
            }
        });
        BRH2.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRH2 = (BRH2.getValue() + "");
            }
        });
        BRH1.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRH1 = (BRH1.getValue() + "");
            }
        });
        BRHS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SBRHS = (BRHS.getValue() + "");
            }
        });
        BGB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SBGB = (BGB.isSelected() + "");
                if (BGB.isSelected() == false) {
                    BHB.setEnabled(false);
                    BHB.setSelected(false);
                    BCB.setEnabled(false);
                    BCB.setSelected(false);
                } else {
                    BHB.setEnabled(true);
                    BCB.setEnabled(true);
                }
            }
        });
        BCB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SBCB = (BCB.isSelected() + "");
            }
        });
        BHB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SBHB = (BHB.isSelected() + "");
            }
        });
        //---------------------------------------
        //             STAGE THREE
        //---------------------------------------
        CRC3.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRC3 = (CRC3.getValue() + "");
            }
        });
        CRC2.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRC2 = (CRC2.getValue() + "");
            }
        });
        CRC1.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRC1 = (CRC1.getValue() + "");
            }
        });
        CRCS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRCS = (CRCS.getValue() + "");
            }
        });
        CRH3.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRH3 = (CRH3.getValue() + "");
            }
        });
        CRH2.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRH2 = (CRH2.getValue() + "");
            }
        });
        CRH1.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRH1 = (CRH1.getValue() + "");
            }
        });
        CRHS.addChangeListener(new ChangeListener() {
            @Override
            public void stateChanged(ChangeEvent e) {
                SCRHS = (CRHS.getValue() + "");
            }
        });
        CGB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SCGB = (CGB.isSelected() + "");
                if (CGB.isSelected() == false) {
                    CHB.setEnabled(false);
                    CHB.setSelected(false);
                    CCB.setEnabled(false);
                    CCB.setSelected(false);
                } else {
                    CHB.setEnabled(true);
                    CCB.setEnabled(true);
                }
            }
        });
        CCB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SCCB = (CCB.isSelected() + "");
            }
        });
        CHB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SCHB = (CHB.isSelected() + "");
            }
        });
        //---------------------------------------
        //             STAGE FOUR
        //---------------------------------------
        DL3.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SDL3 = (DL3.isSelected() + "");
                DL2.setSelected(false);
                DL1.setSelected(false);
            }
        });
        DL2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SDL2 = (DL2.isSelected() + "");
                DL3.setSelected(false);
                DL1.setSelected(false);
            }
        });
        DL1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SDL1 = (DL1.isSelected() + "");
                DL3.setSelected(false);
                DL2.setSelected(false);
            }
        });
        DFB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SDFB = (DFB.isSelected() + "");
            }
        });
/*        DDT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                //String LengthDDT = DDT.getText();
                //int NumDDT = LengthDDT.length();
                SDDT = (DDT.getText());
            }
        });
        DNT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                //String LengthDNT = DNT.getText();
                //int NumDNT = LengthDNT.length();
                SDNT = (DNT.getText());
            }
        });*/
        //------------------------------------------------------------------------
        // JTextFields DATA SETUP IN SUBMIT ACTION LISTENER UNDER "IF" AND "ELSE"
        //------------------------------------------------------------------------

        //ACTION LISTENERS
        EXIT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == EXIT) {
                    System.exit(0);
                }
            }
        });

        NEXTpregame.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == NEXTpregame) {
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.next(cards);
                }
                final long timeInterval = 1000;
                Runnable runnable = new Runnable() {
                    public void run() {
                        while (true) {
                            // ------- code for task to run
                            timer += 1;
                            if ((timer - t) == 0) {
                                System.out.println(15);
                                SandstormTimer.setText("15");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 1) {
                                System.out.println(15);
                                SandstormTimer.setText("15");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 2) {
                                System.out.println(14);
                                SandstormTimer.setText("14");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 3) {
                                System.out.println(13);
                                SandstormTimer.setText("13");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 4) {
                                System.out.println(12);
                                SandstormTimer.setText("12");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 5) {
                                System.out.println(11);
                                SandstormTimer.setText("11");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 6) {
                                System.out.println(10);
                                SandstormTimer.setText("10");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 7) {
                                System.out.println(9);
                                SandstormTimer.setText("9");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 8) {
                                System.out.println(8);
                                SandstormTimer.setText("8");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 9) {
                                System.out.println(7);
                                SandstormTimer.setText("7");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 10) {
                                System.out.println(6);
                                SandstormTimer.setText("6");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 11) {
                                System.out.println(5);
                                SandstormTimer.setText("5");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 12) {
                                System.out.println(4);
                                SandstormTimer.setText("4");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 13) {
                                System.out.println(3);
                                SandstormTimer.setText("3");
                                SandstormTimer.setForeground(Color.BLACK);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 14) {
                                System.out.println(2);
                                SandstormTimer.setText("2");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else if ((timer - t) == 15) {
                                System.out.println(1);
                                SandstormTimer.setText("1");
                                SandstormTimer.setForeground(Color.RED);
                                SandstormTimer.repaint();
                            } else {
                                System.out.println(0);
                                SandstormTimer.setText("0");
                                SandstormTimer.setForeground(Color.RED);
                                NEXTsandstorm.setBackground(Color.ORANGE);
                            }
                            // ------- ends here
                            try {
                                Thread.sleep(timeInterval);
                            } catch (InterruptedException e) {
                                e.printStackTrace();
                            }

                        }
                    }
                };
                Thread thread = new Thread(runnable);
                thread.start();
            }
        });

        NEXTpregame2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == NEXTpregame2) {
                    NEXTsandstorm.setBackground(Color.LIGHT_GRAY);
                    NEXTsandstorm.setForeground(Color.DARK_GRAY);
                    SandstormTimer.setText("15");
                    SandstormTimer.setForeground(Color.BLACK);
                    SandstormTimer.repaint();
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.next(cards);
                }
                t = timer;
            }
        });

        PREVsandstorm.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == PREVsandstorm) {
                    pregame.remove(NEXTpregame);
                    first = false;
                    int x = 0;
                    int y = 0;
                    x = (int) (buttonSize.getWidth() * 4);
                    y = (int) (buttonSize.getHeight() * 9);
                    NEXTpregame2.setBackground(Color.ORANGE);
                    NEXTpregame2.setForeground(Color.BLACK);
                    NEXTpregame2.setFont(buttonFont);
                    NEXTpregame2.setBounds(x, y, 100, 40);
                    pregame.add(NEXTpregame2);
                    NEXTsandstorm.setForeground(Color.DARK_GRAY);
                    NEXTsandstorm.setBackground(Color.LIGHT_GRAY);
                    sandstorm.remove(NEXTsandstorm);
                    x = (int) (buttonSize.getWidth() * 4.34);
                    y = (int) (buttonSize.getHeight() * 10.5);
                    NEXTsandstorm.setBackground(Color.LIGHT_GRAY);
                    NEXTsandstorm.setForeground(Color.DARK_GRAY);
                    NEXTsandstorm.setFont(buttonFont);
                    NEXTsandstorm.setBounds(x, y, 100, 40);
                    sandstorm.add(NEXTsandstorm);
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.first(cards);
                }
            }
        });

        NEXTsandstorm.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == NEXTsandstorm) {
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.next(cards);
                }
            }
        });

        PREVteleop.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == PREVteleop) {
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.previous(cards);
                }
            }
        });

        NEXTteleop.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == NEXTteleop) {
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.next(cards);
                }
            }
        });

        PREVendgame.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == PREVendgame) {
                    CardLayout cl = (CardLayout) (cards.getLayout());
                    cl.previous(cards);
                }
            }
        });

        SUBMIT.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent aActionEvent) {
                //if check to match the code from the question, but not really needed
                if (aActionEvent.getSource() == SUBMIT) {
                    //FALSE
                    SDDT = DDT.getText();
                    SDNT = DNT.getText();
                    //SDDT = (DDT.getText());
                    //SDNT = (DNT.getText());
                    //STRING
                    String FileFormat = ("C:" + '\\' + "Users" + "\\" + UserName + '\\' + "Desktop" + '\\' + "Scout5442-19" + '\\' + "path.txt");
                    //String FileFormat = ("C:" + '\\' + "Users" + "\\" + "arock" + '\\' + "Desktop" + '\\' + "path.txt");

                    String text = "";
                    try {
                        text = new String(Files.readAllBytes(Paths.get(FileFormat)));
                    } catch (IOException d) {
                        d.printStackTrace();
                    }

                    File f = new File("C:" + '\\' + "Users" + "\\" + UserName + '\\' + "Desktop" + '\\' + "Scout5442-19" + '\\' + text + ".xlsx");
                    String f2 = new String("C:" + '\\' + "Users" + "\\" + UserName + '\\' + "Desktop" + '\\' + "Scout5442-19" + '\\' + text + ".xlsx");

                    if (f.exists() && !f.isDirectory()) {

                        try {
                            FileInputStream inputStream = new FileInputStream(new File(f2));
                            Workbook workbook = WorkbookFactory.create(inputStream);

                            Sheet sheet = workbook.getSheetAt(0);

                            Object[][] ScoutingData = {
                                    {SATNS, SBRC3, SBRC2, SBRC1, SBRCS, SBRH3, SBRH2, SBRH1, SBRHS, SBCB, SBHB, SCRC3, SCRC2, SCRC1, SCRCS, SCRH3, SCRH2, SCRH1, SCRHS, SCCB, SCHB, SDL3, SDL2, SDL1, SDFB, SDDT, SDNT},
                            };

                            int rowCount = sheet.getLastRowNum();
                            //int rowCount = 1;

                            for (Object[] aBook : ScoutingData) {
                                Row row = sheet.createRow(rowCount + 1);

                                int columnCount = 0;

                                Cell cell = row.createCell(columnCount);
                                cell.setCellValue(SAMNS);

                                for (Object field : aBook) {
                                    cell = row.createCell(++columnCount);
                                    if (field instanceof String) {
                                        cell.setCellValue((String) field);
                                    } else if (field instanceof Integer) {
                                        cell.setCellValue((Integer) field);
                                    }
                                }

                            }

                            inputStream.close();

                            FileOutputStream outputStream = new FileOutputStream(f2);
                            workbook.write(outputStream);
                            workbook.close();
                            outputStream.close();

                        } catch (IOException | EncryptedDocumentException ex) {
                            ex.printStackTrace();
                        }
                    } else {
                        //SOURCE: https://www.tutorialspoint.com/apache_poi/apache_poi_spreadsheets.htm
                        XSSFWorkbook workbook = new XSSFWorkbook();
                        XSSFSheet spreadsheet = workbook.createSheet(" Sheet1 ");
                        XSSFRow row;

                        Map<String, Object[]> empinfo =
                                new TreeMap<String, Object[]>();
                        empinfo.put("1", new Object[]{"AMNS", "ATNS", "BRC3", "BRC2", "BRC1", "BRCS", "BRH3", "BRH2", "BRH1", "BRHS", "BCB", "BHB", "CRC3", "CRC2", "CRC1", "CRCS", "CRH3", "CRH2", "CRH1", "CRHS", "CCB", "CHB", "DL3", "DL2", "DL1", "DFB", "DDT", "DNT"});
                        empinfo.put("2", new Object[]{SAMNS, SATNS, SBRC3, SBRC2, SBRC1, SBRCS, SBRH3, SBRH2, SBRH1, SBRHS, SBCB, SBHB, SCRC3, SCRC2, SCRC1, SCRCS, SCRH3, SCRH2, SCRH1, SCRHS, SCCB, SCHB, SDL3, SDL2, SDL1, SDFB, SDDT, SDNT});

                        Set<String> keyid = empinfo.keySet();

                        int rowid = 0;
                        for (String key : keyid) {
                            row = spreadsheet.createRow(rowid++);
                            Object[] objectArr = empinfo.get(key);
                            int cellid = 0;

                            for (Object obj : objectArr) {
                                Cell cell = row.createCell(cellid++);
                                cell.setCellValue((String) obj);
                            }
                        }
                        FileOutputStream out = null;
                        try {
                            out = new FileOutputStream(new File(f2));
                        } catch (FileNotFoundException e1) {
                            e1.printStackTrace();
                        }
                        try {
                            workbook.write(out);
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        try {
                            out.close();
                        } catch (IOException e1) {
                            e1.printStackTrace();
                        }
                        System.out.println(f2 + " created written successfully");
                    }

                    //REBOOT

                    //PHASE 1
                    AMNS.setValue(0); //Match Number
                    SAMNS = ("0");
                    ATNS.setValue(0); //Team Number
                    SATNS = ("0");
                    //PHASE 2
                    BRC3.setValue(0); //Sandstorm Rocket 3 Cargo
                    SBRC3 = ("0");
                    BRC2.setValue(0); //Sandstorm Rocket 2 Cargo
                    SBRC2 = ("0");
                    BRC1.setValue(0); //Sandstorm Rocket 1 and Ship Cargo
                    SBRC1 = ("0");
                    BRH3.setValue(0); //Sandstorm Rocket 3 Hatch
                    SBRH3 = ("0");
                    BRH2.setValue(0); //Sandstorm Rocket 2 Hatch
                    SBRH2 = ("0");
                    BRH1.setValue(0); //Sandstorm Rocket 1 and Ship Hatch
                    SBRH1 = ("0");
                    BGB.setSelected(false); //Sandstorm Button to qualify ground pick-up
                    SBGB = ("false");
                    BCB.setSelected(false); //Sandstorm Ground Pickup Cargo
                    BCB.setEnabled(false);
                    SBCB = ("false");
                    BHB.setSelected(false); //Sandstorm Ground Pickup Hatch
                    BHB.setEnabled(false);
                    SBHB = ("false");
                    //PHASE 3
                    CRC3.setValue(0); //Tele-op Rocket 3 Cargo
                    SCRC3 = ("0");
                    CRC2.setValue(0); //Tele-op Rocket 2 Cargo
                    SCRC2 = ("0");
                    CRC1.setValue(0); //Tele-op Rocket 1 and Ship Cargo
                    SCRC1 = ("0");
                    CRH3.setValue(0); //Tele-op Rocket 3 Hatch
                    SCRH3 = ("0");
                    CRH2.setValue(0); //Tele-op Rocket 2 Hatch
                    SCRH2 = ("0");
                    CRH1.setValue(0); //Tele-op Rocket 1 and Ship Hatch
                    SCRH1 = ("0");
                    CGB.setSelected(false); //Tele-op Button to qualify ground pick-up
                    SCGB = ("false");
                    CCB.setSelected(false); //Tele-op Ground Pickup Cargo
                    CCB.setEnabled(false);
                    SCCB = ("false");
                    CHB.setSelected(false); //Tele-op Ground Pickup Hatch
                    CHB.setEnabled(false);
                    SCHB = ("false");
                    //PHASE 4
                    DL3.setSelected(false); //Climb Level 3
                    SDL3 = ("false");
                    DL2.setSelected(false);
                    //Climb Level 2
                    SDL2 = ("false");
                    DL1.setSelected(false); //Climb Level 1
                    SDL1 = ("false");
                    DFB.setSelected(false); //Fouls
                    SDFB = ("false");
                    DDT.setText(""); //Defense Write
                    SDDT = ("---");
                    DNT.setText(""); //Notes Write
                    SDNT = ("---");
                    //NEW ADDITIONS
                    BRCS.setValue(0);
                    SBRCS = "0";
                    BRHS.setValue(0);
                    SBRHS = "0";
                    CRCS.setValue(0);
                    SCRCS = "0";
                    CRHS.setValue(0);
                    SCRHS = "0";

                    pregame.removeAll();
                    //pregame.revalidate();
                    pregame.repaint();
                    first = false;

                    GridLayout dialog = new GridLayout(2, 1);
                    final JDialog d = new JDialog();
                    d.setLayout(dialog);
                    JLabel l = new JLabel("Data has been submitted", SwingConstants.CENTER);
                    d.add(l);
                    d.setSize(200, 100);
                    d.setVisible(true);
                    final JButton b = new JButton("OK");
                    d.add(b);
                    b.addActionListener(new ActionListener() {
                        @Override
                        public void actionPerformed(ActionEvent aActionEvent) {
                            //if check to match the code from the question, but not really needed
                            if (aActionEvent.getSource() == b) {
                                CardLayout cl = (CardLayout) (cards.getLayout());
                                cl.first(cards);
                                d.dispose();
                            }
                        }
                    });

                    int x = 0;
                    int y = 0;

                    JLabel bigPicture = new JLabel(homescreen);
                    bigPicture.setBounds(0, 0, 350, 320);
                    pregame.add(bigPicture);
                    pregame.setBackground(Color.BLACK);
                    Insets insets = pregame.getInsets();
                    pregame.setLayout(null);
                    pregame.setPreferredSize(new Dimension((int) (buttonSize.getWidth() * 5), (int) (buttonSize.getHeight() * 12)));
                    //Picture[insert]
                    //MATCH LABEL
                    x = (int) (buttonSize.getWidth() * 3);
                    y = (int) (buttonSize.getHeight() * 1.8);
                    JLabel PregameMatchNumber = new JLabel("MATCH #", SwingConstants.RIGHT);
                    PregameMatchNumber.setBounds(x, y, 100, 50);
                    PregameMatchNumber.setForeground(Color.ORANGE);
                    PregameMatchNumber.setFont(buttonFont);
                    pregame.add(PregameMatchNumber);
                    //TEAM LABEL
                    x = (int) (buttonSize.getWidth() * 3);
                    y = (int) (buttonSize.getHeight() * 4.3);
                    JLabel PregameTeamNumber = new JLabel("TEAM #", SwingConstants.RIGHT);
                    PregameTeamNumber.setBounds(x, y, 100, 50);
                    PregameTeamNumber.setForeground(Color.ORANGE);
                    PregameTeamNumber.setFont(buttonFont);
                    pregame.add(PregameTeamNumber);
                    //MATCH SPINNER
                    x = (int) (buttonSize.getWidth() * 4);
                    y = (int) (buttonSize.getHeight() * 2);
                    AMNS.setBounds(x, y, 100, 40);
                    AMNS.setForeground(Color.BLACK);
                    pregame.add(AMNS);
                    //TEAM SPINNER
                    x = (int) (buttonSize.getWidth() * 4);
                    y = (int) (buttonSize.getHeight() * 4.5);
                    ATNS.setBounds(x, y, 100, 40);
                    pregame.add(ATNS);
                    x = (int) (buttonSize.getWidth() * 4);
                    y = (int) (buttonSize.getHeight() * 9);
                    NEXTpregame2.setBackground(Color.ORANGE);
                    NEXTpregame2.setForeground(Color.BLACK);
                    NEXTpregame2.setFont(buttonFont);
                    NEXTpregame2.setBounds(x, y, 100, 40);
                    pregame.add(NEXTpregame2);

                }
            }
        });

        //Create the panel that contains the "cards".
        cards = new JPanel(new CardLayout());
        cards.add(pregame);
        cards.add(sandstorm);
        cards.add(teleop);
        cards.add(endgame);

        pane.add(comboBoxPane, BorderLayout.PAGE_START);
        pane.add(cards, BorderLayout.CENTER);

    }

    public void itemStateChanged(ItemEvent evt) {
        CardLayout cl = (CardLayout) (cards.getLayout());
        cl.show(cards, (String) evt.getItem());
    }

    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("5442 Scouting Application 2019");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        //Set Non-Resizable
        frame.setResizable(false);
        //Absolute Layout
        //Create and set up the content pane.
        Main demo = new Main();
        demo.addComponentToPane(frame.getContentPane());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }

    public static void main(String[] args) {

        /* Use an appropriate Look and Feel */
        try {
            //UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
            UIManager.setLookAndFeel("javax.swing.plaf.metal.MetalLookAndFeel");
        } catch (UnsupportedLookAndFeelException ex) {
            ex.printStackTrace();
        } catch (IllegalAccessException ex) {
            ex.printStackTrace();
        } catch (InstantiationException ex) {
            ex.printStackTrace();
        } catch (ClassNotFoundException ex) {
            ex.printStackTrace();
        }
        /* Turn off metal's use of bold fonts */
        UIManager.put("swing.boldMetal", Boolean.FALSE);

        //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }
}