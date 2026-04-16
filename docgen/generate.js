// ============================================================
// SolarPV Pro — Complete Technical Guide & User Manual
// Professional Word Document Generator
// ============================================================

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, UnderlineType,
} = require('docx');
const fs = require('fs');

// ---- Color palette ----
const TEAL = '00695C';
const NAVY = '1A237E';
const WHITE = 'FFFFFF';
const LIGHT_GRAY = 'F5F5F5';
const MID_GRAY = 'E0E0E0';
const CODE_BG = 'F3F4F6';
const ACCENT_BG = 'E8F5E9';
const TABLE_HEADER_BG = '00695C';
const TABLE_ALT_BG = 'F1F8F6';
const DARK_TEXT = '1A1A1A';

// ---- Page settings (A4) ----
const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 1440; // 1 inch
const CONTENT_W = PAGE_W - MARGIN * 2; // 9026 DXA

// ---- Standard border ----
const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

// ============================================================
// HELPERS
// ============================================================

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 200 },
    children: [new TextRun({ text, bold: true, size: 36, color: TEAL, font: 'Arial' })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 4 } },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 28, color: NAVY, font: 'Arial' })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, color: TEAL, font: 'Arial' })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 340 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: DARK_TEXT, ...opts })],
  });
}

function bold(text) {
  return new TextRun({ text, bold: true, size: 22, font: 'Arial', color: DARK_TEXT });
}

function code(text) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    shading: { fill: CODE_BG, type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
    children: [new TextRun({ text, font: 'Courier New', size: 18, color: '1A1A1A' })],
    indent: { left: 360 },
  });
}

function codeBlock(lines) {
  return lines.map(line => code(line));
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: DARK_TEXT })],
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: DARK_TEXT })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function spacer(size = 200) {
  return new Paragraph({ spacing: { before: size, after: 0 }, children: [new TextRun('')] });
}

function mixedPara(runs) {
  return new Paragraph({
    spacing: { after: 120, line: 340 },
    children: runs,
  });
}

function inlineCode(text) {
  return new TextRun({ text, font: 'Courier New', size: 18, shading: { fill: CODE_BG, type: ShadingType.CLEAR }, color: '00695C' });
}

// ---- Table helpers ----
function headerCell(text, colW) {
  return new TableCell({
    borders,
    width: { size: colW, type: WidthType.DXA },
    shading: { fill: TABLE_HEADER_BG, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 20, font: 'Arial' })],
    })],
  });
}

function dataCell(text, colW, isAlt = false, isCenter = false) {
  return new TableCell({
    borders,
    width: { size: colW, type: WidthType.DXA },
    shading: { fill: isAlt ? TABLE_ALT_BG : WHITE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 150, right: 150 },
    children: [new Paragraph({
      alignment: isCenter ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, size: 20, font: 'Arial', color: DARK_TEXT })],
    })],
  });
}

function simpleTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => headerCell(h, colWidths[i])),
      }),
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map((cell, ci) => dataCell(cell, colWidths[ci], ri % 2 === 1)),
        })
      ),
    ],
  });
}

// ============================================================
// DOCUMENT CONTENT
// ============================================================

function buildDoc() {
  const sections_content = [];

  // ====================== COVER PAGE ======================
  const coverPage = [
    spacer(1440),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: 'SolarPV Pro', bold: true, size: 72, color: TEAL, font: 'Arial' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: 'Smart Solar PV System Design &', size: 36, color: NAVY, font: 'Arial', bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: 'Intelligent Load Management', size: 36, color: NAVY, font: 'Arial', bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 800 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL } },
      children: [new TextRun({ text: 'Complete Technical Guide & User Manual', size: 26, color: '555555', font: 'Arial', italics: true })],
    }),
    spacer(400),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      children: [new TextRun({ text: 'EEE3206 Design and Project 2025/2026', size: 24, color: DARK_TEXT, font: 'Arial', bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      children: [new TextRun({ text: 'Department of Electronic Engineering', size: 22, color: '555555', font: 'Arial' })],
    }),
    spacer(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      children: [new TextRun({ text: 'Version 1.0.0  |  April 2026', size: 20, color: '777777', font: 'Arial' })],
    }),
    pageBreak(),
  ];

  // ====================== TABLE OF CONTENTS ======================
  const tocSection = [
    h1('Table of Contents'),
    new TableOfContents('Table of Contents', {
      hyperlink: true,
      headingStyleRange: '1-3',
    }),
    pageBreak(),
  ];

  // ====================== SECTION 1: APP OVERVIEW ======================
  const section1 = [
    h1('1. Application Overview'),
    h2('1.1 Purpose & Background'),
    para('The SolarPV Pro mobile application is a production-ready React Native (Expo) app built for the EEE3206 Design and Project module. It addresses the critical need for accurate, standards-compliant solar PV system design tools that can be used in the field by technicians, students, and solar engineers.'),
    spacer(100),
    para('The application allows users to:'),
    bullet('Automatically size a complete residential solar PV system from user-defined electrical parameters'),
    bullet('Apply IEC/NEC engineering standards for cable sizing, protection devices, and component sizing'),
    bullet('Intelligently manage electrical loads based on battery State-of-Charge (SOC) and solar generation'),
    bullet('Visualize the system architecture through interactive SVG diagrams and animated data charts'),
    bullet('Save and compare multiple system designs for offline use without internet connectivity'),
    spacer(100),
    para('Correct design and management of residential solar PV systems is critical to ensure safety, efficiency, reliability, and long system lifespan. Many solar PV systems installed in residential environments suffer from poor performance, frequent inverter overloads, excessive battery discharge, and component failures due to incorrect system sizing and lack of proper load management. This application directly addresses these problems.'),
    spacer(200),
    h2('1.2 Technology Stack'),
    para('The following table lists all technologies used in building the application:'),
    spacer(100),
    simpleTable(
      ['Technology', 'Purpose', 'Version'],
      [
        ['React Native (Expo)', 'Cross-platform mobile framework', 'Expo SDK 52'],
        ['TypeScript (strict)', 'Type-safe development language', '5.3+'],
        ['Zustand', 'Global state management', '5.0+'],
        ['React Navigation', 'Screen navigation (Stack + Bottom Tabs)', '6.x'],
        ['React Native SVG', 'Custom charts and system diagrams', '15.8'],
        ['React Native Reanimated', 'Smooth animations and transitions', '3.16'],
        ['React Native Gesture Handler', 'Touch and gesture interactions', '2.20'],
        ['Zod', 'Form and data schema validation', '3.24'],
        ['AsyncStorage', 'Offline data persistence (local storage)', '2.1'],
        ['Expo Linear Gradient', 'Glassmorphism UI gradient effects', '14.0'],
        ['@expo/vector-icons', 'Ionicons icon library', '14.0'],
      ],
      [3200, 3600, 2200]
    ),
    pageBreak(),
  ];

  // ====================== SECTION 2: FILE STRUCTURE ======================
  const section2 = [
    h1('2. Project File Structure'),
    para('The project follows a clean, modular architecture with clear separation of concerns. Every file has a single responsibility.'),
    spacer(100),
    ...codeBlock([
      'SolarPVApp/',
      '├── App.tsx              Root entry point',
      '├── app.json             Expo configuration file',
      '├── babel.config.js      Babel transpiler config',
      '├── package.json         Dependencies and npm scripts',
      '├── tsconfig.json        TypeScript strict mode config',
      '└── src/',
      '    ├── types/',
      '    │   └── index.ts     All TypeScript interfaces (30+ types)',
      '    ├── constants/',
      '    │   ├── theme.ts     Design tokens (colors, spacing)',
      '    │   └── solar.ts     Engineering constants (IEC standards)',
      '    ├── utils/',
      '    │   ├── solarCalculations.ts   PV design engine',
      '    │   ├── loadManagement.ts      Load priority algorithm',
      '    │   └── validation.ts          Zod form schemas',
      '    ├── store/',
      '    │   └── useStore.ts  Zustand global state',
      '    ├── hooks/',
      '    │   └── useSystemDesign.ts     Custom React hooks',
      '    ├── services/',
      '    │   └── storage.ts   AsyncStorage wrapper',
      '    ├── components/',
      '    │   ├── ui/          7 reusable UI primitives',
      '    │   ├── charts/      4 custom SVG chart components',
      '    │   └── 3d/          Interactive system diagram',
      '    ├── screens/         8 full application screens',
      '    └── navigation/',
      '        └── AppNavigator.tsx  Stack + Tab navigation',
    ]),
    spacer(200),
    h2('2.1 Configuration Files'),
    simpleTable(
      ['File', 'Purpose'],
      [
        ['App.tsx', 'Root component. Wraps the entire app in GestureHandlerRootView (required by React Native Gesture Handler) and renders the AppNavigator.'],
        ['app.json', 'Expo SDK configuration. Defines app name, slug, version, icon, splash screen color, orientation (portrait), and platform-specific settings.'],
        ['babel.config.js', 'Babel transpiler configuration. Extends babel-preset-expo and adds the react-native-reanimated/plugin (MUST be the last plugin in the list).'],
        ['package.json', 'NPM package manifest listing all 17 production dependencies and 3 dev dependencies with their pinned versions.'],
        ['tsconfig.json', 'TypeScript configuration with strict mode enabled. Sets up path aliases (@/* → src/*) for clean imports.'],
      ],
      [2200, 6826]
    ),
    spacer(200),
    h2('2.2 Source Code Files'),
    h3('types/index.ts'),
    para('Defines all TypeScript interfaces used across the entire application. Key types include:'),
    bullet('SystemInputs — all user-entered parameters for a design'),
    bullet('DesignResults — calculated output (PV, battery, inverter, cables, protection)'),
    bullet('LoadItem — a single electrical load with priority, power, and schedule'),
    bullet('LoadManagementState — real-time load decisions with reasons'),
    bullet('SimulationDataPoint — hourly data point for 24h simulation'),
    bullet('SavedDesign — a named saved design with inputs and results'),
    bullet('RootStackParamList / TabParamList — Navigation type safety'),
    spacer(100),
    h3('constants/theme.ts'),
    para('Single source of truth for all UI design tokens. Exports:'),
    bullet('Colors — full palette including glassmorphism glass/glassBorder values, gradient arrays, status colors'),
    bullet('Spacing — xs(4), sm(8), md(16), lg(24), xl(32), xxl(48) in pixels'),
    bullet('BorderRadius — sm(8), md(12), lg(16), xl(24), full(999)'),
    bullet('FontSize — xs(10) through display(48)'),
    bullet('FontWeight — regular, medium, semibold, bold, heavy'),
    bullet('Shadows — small, medium, large, and glow(color) factory function'),
    spacer(100),
    h3('constants/solar.ts'),
    para('All electrical engineering constants per IEC/NEC standards:'),
    bullet('STANDARD_CABLE_SIZES — [1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120] mm²'),
    bullet('COPPER_RESISTIVITY — 0.0175 Ω·mm²/m at 20°C'),
    bullet('STANDARD_BREAKER_RATINGS — [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200] A'),
    bullet('STANDARD_INVERTER_SIZES — standard commercial inverter ratings in Watts'),
    bullet('CC_SAFETY_FACTOR — 1.25 (charge controller safety multiplier)'),
    bullet('INVERTER_SURGE_FACTOR — 1.25 (motor startup surge allowance)'),
    bullet('TEMP_COEFFICIENTS — Voc (-0.3%/°C), Isc (+0.05%/°C), Pmax (-0.4%/°C)'),
    bullet('PEAK_SUN_HOURS_BY_REGION — lookup table for 11 global regions'),
    bullet('DEFAULT_PV_MODULE — 400Wp typical module parameters'),
    bullet('SAMPLE_LOADS — 8 pre-defined demonstration loads'),
    pageBreak(),
  ];

  // ====================== SECTION 3: CALCULATION ENGINE ======================
  const section3 = [
    h1('3. Engineering Calculation Engine'),
    para('The file src/utils/solarCalculations.ts implements six complete engineering calculation modules. All formulas are based on IEC 62548, NEC Article 690, and IEC 60364-7-712 standards.'),
    spacer(100),
    h2('3.1 PV Array Sizing'),
    para('The PV array sizing module calculates the number of solar panels required and their series/parallel configuration.'),
    spacer(100),
    h3('Primary Formula'),
    code('Total Energy Required = Daily Energy Consumption / System Efficiency'),
    code('Panel Energy per Day  = Panel Power × PSH × η_soiling × η_mismatch'),
    code('Number of Panels      = CEIL(Total Energy Required / Panel Energy per Day)'),
    spacer(100),
    h3('Temperature Corrections'),
    para('All electrical parameters are corrected for operating temperature before use:'),
    code('T_cell = T_ambient + 25°C              (NOCT approximation)'),
    code('ΔT     = T_cell - 25°C                 (delta from STC)'),
    code('Voc_corrected = Voc × (1 + αVoc × ΔT)  αVoc = -0.003 per °C'),
    code('Isc_corrected = Isc × (1 + αIsc × ΔT)  αIsc = +0.0005 per °C'),
    spacer(100),
    h3('Series/Parallel Configuration'),
    code('Panels in Series   = CEIL(Battery Voltage / Module Vmp)'),
    code('Panels in Parallel = CEIL(Total Panels / Series Count)'),
    code('Actual Total       = Series × Parallel  (adjusted to complete strings)'),
    code('Array Voc          = Series Count × Voc_corrected'),
    code('Array Isc          = Parallel Count × Isc_corrected'),
    spacer(100),
    h3('Derating Factors Used'),
    simpleTable(
      ['Factor', 'Value', 'Reason'],
      [
        ['Dust & Soiling', '0.95 (5% loss)', 'Accumulated dirt on panel surface'],
        ['Module Mismatch', '0.98 (2% loss)', 'Manufacturing tolerance variation'],
        ['Battery Charging', '0.85 (15% loss)', 'Charging efficiency of lead-acid'],
        ['System Efficiency', 'User-defined (default 85%)', 'Combined wiring, conversion losses'],
      ],
      [2200, 2500, 4326]
    ),
    spacer(200),
    h2('3.2 Battery Bank Sizing'),
    para('The battery sizing module determines the required Ah capacity and physical battery configuration.'),
    code('C_required (Ah) = (E_daily × Autonomy_Days) / (V_battery × DoD × η_battery)'),
    code('Usable Capacity  = C_required × DoD'),
    code('Batteries in Series   = V_system / V_single_battery  (12V cells assumed)'),
    code('Batteries in Parallel = CEIL(C_required / C_single_battery)  (200Ah cells)'),
    code('Total Batteries       = Series × Parallel'),
    code('Bank Energy (Wh)      = Total_Ah × V_system'),
    spacer(100),
    para('Depth of Discharge (DoD) guidance:'),
    simpleTable(
      ['Battery Type', 'Recommended DoD', 'Cycle Life at DoD'],
      [
        ['Lead-Acid (flooded)', '50%', '~500 cycles'],
        ['AGM / GEL', '60%', '~600 cycles'],
        ['Lithium (LiFePO4)', '80–90%', '2000+ cycles'],
      ],
      [2800, 2800, 3426]
    ),
    spacer(200),
    h2('3.3 Inverter Sizing'),
    code('Total Load Power  = Σ (Load_Power × Quantity)  for all loads'),
    code('Surge Power       = Total Load Power × 1.25    (startup inrush)'),
    code('Recommended Min   = Surge Power (rounded up)'),
    code('Selected Rating   = next standard size ≥ Recommended Min'),
    para('Standard inverter sizes checked: 300W, 500W, 800W, 1000W, 1500W, 2000W, 3000W, 5000W, 8000W, 10000W'),
    spacer(200),
    h2('3.4 Charge Controller Sizing'),
    para('The controller type (MPPT vs PWM) is automatically selected based on system parameters:'),
    code('MPPT selected when: Array Power > 500W  OR  Battery Voltage >= 48V'),
    code('PWM used when:      Array Power <= 500W AND Battery Voltage < 48V'),
    spacer(100),
    para('Rating calculations differ by type:'),
    code('MPPT Rating = (Total Array Power / Battery Voltage) × 1.25  [A]'),
    code('PWM Rating  = Isc_corrected × Parallel Strings × 1.25       [A]'),
    code('Max Input Voltage = Series Panels × Voc_cold                 [V]'),
    para('Voc_cold uses minimum expected temperature (0°C) for worst-case open-circuit voltage — critical for controller protection.'),
    spacer(100),
    para('Standard controller sizes: 10A, 20A, 30A, 40A, 50A, 60A, 80A, 100A, 150A, 200A'),
    spacer(200),
    h2('3.5 Cable Sizing — IEC Voltage Drop Method'),
    para('Four separate cable runs are calculated. The minimum cross-sectional area is determined by the allowable voltage drop, then rounded up to the next standard IEC cable size.'),
    spacer(100),
    h3('Core Formula'),
    code('A_min (mm²) = (ρ × 2L × I) / (V_system × Δv_max)'),
    spacer(100),
    para('Where:'),
    bullet('ρ = 0.0175 Ω·mm²/m (copper resistivity at 20°C)'),
    bullet('2L = two-way cable length (current flows both directions)'),
    bullet('I = maximum circuit current in Amperes'),
    bullet('V_system = nominal circuit voltage'),
    bullet('Δv_max = maximum allowable fractional voltage drop'),
    spacer(100),
    simpleTable(
      ['Circuit', 'Voltage', 'Max Drop', 'Current Basis'],
      [
        ['PV Array → Controller', 'Battery V', '3% (DC)', 'Isc × 1.25'],
        ['Controller → Battery', 'Battery V', '3% (DC)', 'Array Power / V_bat × 1.25'],
        ['Battery → Inverter', 'Battery V', '3% (DC)', 'Inverter Power / (V_bat × η_inv) × 1.25'],
        ['Inverter → Load Panel', '230V AC', '5% (AC)', 'Inverter Power / 230 × 1.25'],
      ],
      [2600, 1600, 1600, 3226]
    ),
    spacer(100),
    para('Standard cable sizes selected from: 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120 mm²'),
    spacer(200),
    h2('3.6 Protection Devices — NEC 690.9'),
    para('Protection ratings are calculated per NEC Article 690.9 and IEC 60269 standards:'),
    code('PV Array Fuse    = CEIL(Isc_array × 1.56)  [NEC 690.9]'),
    code('Battery Breaker  = CEIL((P_inverter / (V_bat × 0.90)) × 1.25)'),
    code('AC Inverter OCPD = CEIL((P_inverter / 230V) × 1.25)'),
    code('Load Panel OCPD  = same as AC Inverter OCPD'),
    spacer(100),
    para('Standard breaker ratings: 6A, 10A, 16A, 20A, 25A, 32A, 40A, 50A, 63A, 80A, 100A, 125A, 160A, 200A'),
    para('Additional requirements: surge protection device (SPD) and earthing/grounding are always flagged as required by the application.'),
    pageBreak(),
  ];

  // ====================== SECTION 4: LOAD MANAGEMENT ======================
  const section4 = [
    h1('4. Intelligent Load Management Algorithm'),
    para('The load management engine (src/utils/loadManagement.ts) implements a rule-based decision system that determines which electrical loads should be energised at any given time, based on battery SOC, solar generation, and time of day.'),
    spacer(100),
    h2('4.1 Load Priority Classification'),
    simpleTable(
      ['Priority', 'Description', 'Examples', 'Kept ON When'],
      [
        ['Critical', 'Essential for safety and health', 'Lights, fridge, medical equipment, security', 'Always (if power available)'],
        ['Important', 'High value but deferrable', 'TV, laptop, phone charger, router', 'SOC > 40%'],
        ['Non-Critical', 'Convenience loads', 'Iron, microwave, fans, entertainment', 'SOC > 60% + solar generating'],
      ],
      [1800, 2200, 2600, 2426]
    ),
    spacer(200),
    h2('4.2 SOC Threshold Decision Rules'),
    para('The algorithm uses four SOC zones. Each zone activates different load shedding rules:'),
    simpleTable(
      ['Zone', 'SOC Range', 'System Status', 'Loads Permitted'],
      [
        ['Critical', '0–20%', 'CRITICAL (Red)', 'Critical loads only. All others shed.'],
        ['Warning', '21–40%', 'WARNING (Yellow)', 'Critical + Important loads. Non-critical shed.'],
        ['Normal', '41–60%', 'NORMAL', 'All loads, but non-critical requires solar generation.'],
        ['Optimal', '61–100%', 'OPTIMAL (Green)', 'All loads freely permitted.'],
      ],
      [1600, 1600, 2200, 3626]
    ),
    spacer(100),
    para('Additional time-of-day rule: Non-critical loads after 18:00 (evening) are deferred if SOC is below 70%, regardless of the SOC zone. This preserves battery reserve for night-time essential loads.'),
    spacer(200),
    h2('4.3 Solar Generation Model'),
    para('Solar generation is modelled using a Gaussian (bell) curve peaking at solar noon. This provides a realistic approximation of irradiance throughout the day.'),
    code('G(h) = G_peak × exp(-0.5 × ((h - 12) / σ)²)'),
    code('σ = 3  (standard deviation, controls curve width)'),
    spacer(100),
    simpleTable(
      ['Scenario', 'Generation', 'Description'],
      [
        ['Sunny', '100% of G_peak', 'Clear sky, full array output'],
        ['Cloudy', '35% of G_peak', 'Overcast, scattered irradiance'],
        ['Night', '0W (hours 0-5, 19-23)', 'No solar generation'],
      ],
      [2000, 2500, 4526]
    ),
    spacer(200),
    h2('4.4 24-Hour Simulation Algorithm'),
    para('The simulate24Hours() function generates hourly system data for chart display:'),
    code('FOR each hour h from 0 to 23:'),
    code('  solarGen(h)   = G_peak × gaussian(h, scenario)'),
    code('  consumption   = Σ active_load_power for loads active at hour h'),
    code('  netPower      = solarGen - consumption'),
    code('  ΔEnergy       = netPower  (Wh per hour)'),
    code('  SOC(h)        = SOC(h-1) + (ΔEnergy / Battery_Capacity_Wh × 100)'),
    code('  SOC(h)        = CLAMP(SOC(h), 0, 100)'),
    spacer(100),
    para('Load activity is modelled with realistic time patterns:'),
    bullet('Lights: active 18:00–23:59 and 00:00–06:00 (night hours)'),
    bullet('Refrigerator: active 24 hours continuously'),
    bullet('Fans: active 10:00–22:00 (daytime and evening)'),
    bullet('TV, laptop, chargers: active 07:00–22:00 (waking hours)'),
    pageBreak(),
  ];

  // ====================== SECTION 5: SCREEN GUIDE ======================
  const section5 = [
    h1('5. Screen-by-Screen User Guide'),
    h2('5.1 Home Screen'),
    para('The entry point of the application. Displayed on launch and accessible via the "Home" tab.'),
    spacer(100),
    h3('Header Banner'),
    bullet('Displays app name "SolarPV Pro" with teal-to-green gradient'),
    bullet('Three quick-stat items: Daily Energy (Wh), Number of Loads, and Number of Panels'),
    bullet('Stats update automatically as you add loads and run calculations'),
    h3('Quick Action Cards'),
    bullet('New Design — opens the 6-step Design Wizard'),
    bullet('Load Manager — opens the intelligent load control screen'),
    bullet('3D View — opens the interactive system diagram screen'),
    bullet('Saved Designs — opens your saved design library'),
    h3('Current Design Panel'),
    bullet('Appears only after a calculation has been run'),
    bullet('Shows PV array power, battery Ah, inverter W, and controller A at a glance'),
    bullet('Tap "View Details" to navigate to the full Results screen'),
    h3('Get Started CTA'),
    bullet('Shown when no design exists yet'),
    bullet('"Start Design Wizard" button launches directly into Step 1'),
    spacer(200),
    h2('5.2 Design Wizard — 6 Steps'),
    para('A guided multi-step form with animated transitions between steps. Validation is performed at each step before allowing advancement.'),
    spacer(100),
    simpleTable(
      ['Step', 'Title', 'Fields'],
      [
        ['1', 'Energy & Battery', 'Daily energy (Wh/day), Battery voltage (12/24/48V), Autonomy days, Depth of discharge (%)'],
        ['2', 'PV Module Specs', 'Rated Power (Wp), Open-circuit Voltage (Voc), Short-circuit Current (Isc), Vmp, Imp'],
        ['3', 'Site Conditions', 'Peak sun hours (with regional quick-select), Ambient temperature (°C), System efficiency (%)'],
        ['4', 'Inverter Setup', 'Inverter type (Pure Sine / Modified Sine), Rated power (W)'],
        ['5', 'Cable Lengths', 'PV→Controller (m), Controller→Battery (m), Battery→Inverter (m), Inverter→Load Panel (m)'],
        ['6', 'Load Schedule', 'Add loads: name, power (W), quantity, hours/day, priority (Critical/Important/Non-Critical)'],
      ],
      [800, 2000, 6226]
    ),
    spacer(100),
    h3('Navigation Controls'),
    bullet('Back button: returns to previous step or exits wizard'),
    bullet('Next button: validates current step then advances'),
    bullet('Step dots at top: visual indicator of current position'),
    bullet('Progress bar: animated fill showing percentage complete'),
    bullet('"Calculate System" button on Step 6: triggers full system calculation and navigates to Results'),
    spacer(100),
    h3('Load Entry (Step 6)'),
    bullet('Fill in appliance name, power (W), quantity, and daily hours'),
    bullet('Select priority using the segmented control (Critical / Important / Non-Critical)'),
    bullet('Tap "Add Load" to append to the load list'),
    bullet('Daily energy (Wh) for each load is shown automatically'),
    bullet('Tap the trash icon to remove a load'),
    bullet('At least one load must be added before calculation is permitted'),
    spacer(200),
    h2('5.3 Results Screen'),
    para('Displays all calculated system parameters in six collapsible cards with animated reveal. Accessible via the wizard completion or from the Home screen.'),
    spacer(100),
    h3('Result Cards'),
    bullet('PV Array Card — Total panel count, S×P configuration, Array Voc, Array Isc, daily energy production'),
    bullet('Battery Bank Card — Required Ah, total bank Ah, number of batteries, bank energy (Wh), usable capacity'),
    bullet('Inverter Card — Selected standard rating, surge power, total connected load'),
    bullet('Charge Controller Card — Rating (A), MPPT or PWM type, max input voltage, max input current'),
    bullet('Cable Sizes Card — mm² size and actual voltage drop % for all four cable runs'),
    bullet('Protection Devices Card — PV fuse, battery breaker, inverter breaker, load panel breaker ratings'),
    spacer(100),
    h3('Actions'),
    bullet('Bookmark icon (top right): saves the current design with today\'s date as name'),
    bullet('"View 3D Diagram" button: navigates to Visualization screen'),
    bullet('"Load Manager" button: navigates to Load Management screen'),
    spacer(200),
    h2('5.4 Dashboard Screen'),
    para('The central data visualization hub. Accessible via the "Dashboard" tab.'),
    spacer(100),
    h3('Controls'),
    bullet('Weather Scenario segmented control: Sunny / Cloudy / Night — instantly re-runs the 24h simulation'),
    h3('Metric Cards Row'),
    bullet('Solar Array (Wp total), Daily Load (Wh/day), Average Battery SOC (%)'),
    h3('Energy Flow Chart'),
    bullet('24-hour area line chart showing solar generation (yellow) vs load consumption (red)'),
    bullet('SVG-rendered with gradient fills under each curve'),
    bullet('Grid lines and hour/value axis labels'),
    h3('Battery Section'),
    bullet('Circular arc gauge showing average SOC — color changes: green (optimal), yellow (warning), red (critical)'),
    bullet('Stats panel: Max SOC, Min SOC, Capacity (Ah), Energy (kWh)'),
    h3('SOC Timeline'),
    bullet('24-hour SOC tracking chart with warning zone shading (red below 20%, yellow 20–40%)'),
    bullet('Dashed threshold lines at 20% and 40%'),
    h3('Load Distribution Donut Chart'),
    bullet('Shows energy split (Wh/day) by priority: Critical (red), Important (blue), Non-Critical (grey)'),
    spacer(200),
    h2('5.5 Load Management Screen'),
    para('Real-time load control simulator. Accessible from the "Loads" tab or from the Results screen.'),
    spacer(100),
    h3('Status Banner'),
    bullet('Shows current solar generation (W), active total load (W), and available power (W)'),
    bullet('Banner color changes based on system status: green/yellow/red gradient'),
    h3('Simulation Controls'),
    bullet('Weather scenario: Sunny / Cloudy / Night'),
    bullet('Battery SOC: tap a preset value (0%, 20%, 40%, 60%, 80%, 100%)'),
    bullet('Time of day: tap any hour (0–23) on the horizontal timeline'),
    bullet('All controls immediately trigger recalculation of load decisions'),
    h3('Battery Gauge'),
    bullet('Circular SOC gauge updating in real-time as you adjust controls'),
    h3('Load Decisions List'),
    bullet('Every configured load shown with name, power (W), priority badge, and ON/OFF status'),
    bullet('Color-coded status indicator (green bar = ON, red bar = OFF)'),
    bullet('Explanation text shown under each load explaining why it is on or off'),
    spacer(200),
    h2('5.6 System Visualization Screen'),
    para('Interactive SVG system diagram showing all components and energy flow.'),
    spacer(100),
    h3('System Diagram'),
    bullet('Animated sun with rays (shown when solar generation is ON)'),
    bullet('Solar panel array with grid cell pattern'),
    bullet('Charge controller with three LED indicator dots'),
    bullet('Battery bank with real-time fill level showing SOC percentage'),
    bullet('Inverter with DC→AC label'),
    bullet('House/load block'),
    bullet('Animated dashed flow lines between all components with directional arrows'),
    h3('Tap Interactions'),
    bullet('Tap any component to open a detail popup showing key specifications from your calculated results'),
    bullet('Tap anywhere outside the popup to dismiss it'),
    h3('Controls'),
    bullet('Solar ON/OFF toggle: shows/hides the sun and changes flow line colors'),
    bullet('SOC level selector: 10%, 30%, 50%, 70%, 90% — changes battery fill and glow color'),
    bullet('Component Summary card at bottom lists all four component types with ratings'),
    spacer(200),
    h2('5.7 Saved Designs Screen'),
    bullet('Lists all designs saved via the bookmark icon on the Results screen'),
    bullet('Each card shows design name, date/time, and four key figures (Array, Battery, Inverter, Panels)'),
    bullet('Tap "Load Design" to restore a saved design — navigates to Results screen'),
    bullet('Tap the trash icon to delete a design (with confirmation dialog)'),
    bullet('Empty state shown with call-to-action when no designs are saved'),
    spacer(200),
    h2('5.8 Settings Screen'),
    bullet('Displays app name, version (1.0.0), and description'),
    bullet('Shows count of saved designs'),
    bullet('"Reset Inputs" button clears all wizard inputs back to defaults (saved designs are not affected)'),
    bullet('Engineering standards reference section listing all IEC/NEC standards used'),
    pageBreak(),
  ];

  // ====================== SECTION 6: SETUP GUIDE ======================
  const section6 = [
    h1('6. Complete Setup & Installation Guide'),
    h2('6.1 Prerequisites'),
    para('Install the following software before proceeding:'),
    spacer(100),
    simpleTable(
      ['Software', 'Version', 'Download', 'Required?'],
      [
        ['Node.js', 'v20 LTS (recommended)', 'nodejs.org', 'Yes'],
        ['npm', 'Comes with Node.js', 'auto-installed', 'Yes'],
        ['Git', 'Any recent version', 'git-scm.com', 'Optional'],
        ['Expo Go (mobile)', 'Latest from App Store / Play Store', 'Search "Expo Go"', 'For phone testing'],
        ['Android Studio', 'Latest stable', 'developer.android.com', 'For emulator/APK'],
        ['Xcode', 'Latest (Mac only)', 'Mac App Store', 'For iOS (Mac only)'],
      ],
      [2200, 2200, 2500, 2126]
    ),
    spacer(100),
    para('IMPORTANT: Node.js v24 has a known incompatibility with the analytics module in Expo 50\'s CLI. This project has been upgraded to Expo SDK 52 which resolves this. However, using Node.js v20 LTS is still recommended for maximum compatibility.'),
    spacer(100),
    para('Verify your Node.js version:'),
    code('node --version'),
    code('npm --version'),
    spacer(200),
    h2('6.2 Step-by-Step Installation'),
    h3('Step 1: Open Terminal / PowerShell'),
    para('On Windows, right-click the Start button and select "Windows PowerShell" or "Terminal".'),
    spacer(100),
    h3('Step 2: Navigate to the Project Directory'),
    code('cd C:\\Users\\TechCharities\\Documents\\final_project\\SolarPVApp'),
    spacer(100),
    h3('Step 3: Install Dependencies'),
    code('npm install --legacy-peer-deps'),
    para('This installs all 17 production dependencies. The --legacy-peer-deps flag resolves minor peer dependency conflicts. Expected output: "added ~500 packages".'),
    spacer(100),
    h3('Step 4: Start the Development Server'),
    code('npx expo start'),
    para('The Metro bundler starts and displays a QR code in the terminal. Keep this terminal window open.'),
    spacer(100),
    h3('Step 5: Run on Your Device'),
    spacer(100),
    para('Option A — Physical Android/iOS phone (easiest, no setup required):'),
    numbered('Install the Expo Go app from your device\'s app store'),
    numbered('Ensure your phone and PC are on the same WiFi network'),
    numbered('Open Expo Go, tap "Scan QR code"'),
    numbered('Scan the QR code displayed in the terminal'),
    numbered('The SolarPV Pro app loads and runs on your phone'),
    spacer(100),
    para('Option B — Android Emulator (via Android Studio):'),
    numbered('Open Android Studio → Tools → Device Manager'),
    numbered('Create or start an existing AVD (Android Virtual Device)'),
    numbered('In the Expo terminal, press "a"'),
    numbered('Expo installs and launches the app on the emulator automatically'),
    spacer(100),
    para('Option C — iOS Simulator (Mac only):'),
    numbered('Open Xcode → Open Developer Tool → Simulator'),
    numbered('In the Expo terminal, press "i"'),
    numbered('App launches in the iOS simulator'),
    pageBreak(),
  ];

  // ====================== SECTION 7: APK BUILD ======================
  const section7 = [
    h1('7. Building an APK for Distribution'),
    para('An APK (Android Package) file allows you to install the app on any Android device without needing a development server running. There are two methods: cloud build (EAS) and local build.'),
    spacer(100),
    h2('7.1 Method 1: EAS Cloud Build (Recommended)'),
    para('Expo Application Services (EAS) builds the APK on Expo\'s cloud servers. No Android Studio or Java required on your machine. A free Expo account is all you need.'),
    spacer(100),
    h3('Step 1: Install EAS CLI'),
    code('npm install -g eas-cli'),
    spacer(100),
    h3('Step 2: Create a Free Expo Account'),
    bullet('Visit https://expo.dev and click "Sign Up"'),
    bullet('Verify your email address'),
    spacer(100),
    h3('Step 3: Log In via Terminal'),
    code('eas login'),
    para('Enter your expo.dev email and password when prompted.'),
    spacer(100),
    h3('Step 4: Configure EAS in Your Project'),
    code('cd C:\\Users\\TechCharities\\Documents\\final_project\\SolarPVApp'),
    code('eas build:configure'),
    para('When prompted, select Android. This creates an eas.json file in your project.'),
    spacer(100),
    h3('Step 5: Configure eas.json for APK Output'),
    para('Open eas.json and replace its content with:'),
    code('{'),
    code('  "cli": {'),
    code('    "version": ">= 12.0.0"'),
    code('  },'),
    code('  "build": {'),
    code('    "preview": {'),
    code('      "android": {'),
    code('        "buildType": "apk"'),
    code('      }'),
    code('    },'),
    code('    "production": {'),
    code('      "android": {'),
    code('        "buildType": "apk"'),
    code('      }'),
    code('    }'),
    code('  }'),
    code('}'),
    para('IMPORTANT: The default EAS build produces an AAB (Android App Bundle) for Play Store submission. The "buildType": "apk" setting forces it to produce a directly-installable APK file.'),
    spacer(100),
    h3('Step 6: Start the Cloud Build'),
    code('eas build -p android --profile preview'),
    para('This uploads your project source to Expo\'s build servers. Expected duration: 5–15 minutes. You will receive a terminal link and email notification when complete.'),
    spacer(100),
    h3('Step 7: Download the APK'),
    bullet('Visit https://expo.dev/accounts/[your-username]/projects/solar-pv-app/builds'),
    bullet('Click the completed build'),
    bullet('Click "Download" to save the .apk file to your computer'),
    spacer(200),
    h2('7.2 Method 2: Local APK Build'),
    para('Builds the APK locally using Android Studio and Gradle. Requires Android Studio to be installed and configured.'),
    spacer(100),
    h3('Prerequisites'),
    bullet('Android Studio installed with Android SDK'),
    bullet('Java Development Kit (JDK) 17 — comes with Android Studio'),
    bullet('ANDROID_HOME environment variable set (usually auto-configured by Android Studio)'),
    spacer(100),
    h3('Step 1: Generate Native Android Project'),
    code('cd C:\\Users\\TechCharities\\Documents\\final_project\\SolarPVApp'),
    code('npx expo prebuild --platform android'),
    para('This creates an /android directory containing a complete Gradle-based Android project. Do not edit files in this directory manually.'),
    spacer(100),
    h3('Step 2: Build the APK'),
    para('For a debug APK (for testing and sharing):'),
    code('cd android'),
    code('gradlew assembleDebug'),
    spacer(100),
    para('For a release APK (optimised, requires signing):'),
    code('gradlew assembleRelease'),
    spacer(100),
    h3('Step 3: Locate the APK File'),
    para('Debug APK location:'),
    code('android\\app\\build\\outputs\\apk\\debug\\app-debug.apk'),
    spacer(100),
    para('Release APK location:'),
    code('android\\app\\build\\outputs\\apk\\release\\app-release.apk'),
    spacer(200),
    h2('7.3 Installing the APK on Android Devices'),
    h3('Enable Unknown Sources'),
    numbered('Open Settings on your Android phone'),
    numbered('Go to Security (or Privacy on some devices)'),
    numbered('Enable "Install Unknown Apps" or "Unknown Sources"'),
    numbered('On Android 8+: go to Settings → Apps → Special App Access → Install Unknown Apps → enable for your file manager'),
    spacer(100),
    h3('Transfer and Install'),
    para('Transfer the APK to your phone using any of these methods:'),
    simpleTable(
      ['Transfer Method', 'How To', 'Notes'],
      [
        ['USB Cable', 'Connect phone to PC, copy APK to Downloads folder', 'Fastest, most reliable'],
        ['Google Drive', 'Upload APK → share link → open on phone browser', 'Works across networks'],
        ['WhatsApp', 'Send APK to yourself as a Document (not media)', 'File must be < 100MB'],
        ['Email', 'Attach APK to email, open on phone', 'Check attachment size limits'],
        ['Bluetooth', 'Send file via Bluetooth from PC to phone', 'Slow but works offline'],
      ],
      [2000, 4000, 3026]
    ),
    spacer(100),
    para('After transferring:'),
    numbered('Open the APK file on your phone using a file manager'),
    numbered('Tap "Install" when prompted'),
    numbered('Wait for installation to complete (10–30 seconds)'),
    numbered('Tap "Open" — SolarPV Pro is now installed on your device'),
    spacer(200),
    h2('7.4 Sharing the APK with Others'),
    simpleTable(
      ['Platform', 'Method', 'Best For'],
      [
        ['Google Drive', 'Upload APK → Get Link → share URL', 'Sharing with multiple people'],
        ['WhatsApp Group', 'Send as Document to your group', 'Fast team distribution'],
        ['QR Code', 'Upload to Drive → paste link at qr-code-generator.com', 'In-person sharing'],
        ['Telegram', 'Send as file (no compression)', 'Large files, community sharing'],
        ['GitHub Releases', 'Create a release and attach APK', 'Open source distribution'],
      ],
      [2000, 3500, 3526]
    ),
    pageBreak(),
  ];

  // ====================== SECTION 8: SAMPLE TEST CASE ======================
  const section8 = [
    h1('8. Sample Test Case with Manual Verification'),
    para('The following worked example demonstrates how to use the app and verify the results against manual calculations.'),
    spacer(100),
    h2('8.1 Test Case Inputs'),
    simpleTable(
      ['Parameter', 'Value'],
      [
        ['Load 1', '4× LED Lights, 10W each, 6 hours/day (Critical)'],
        ['Load 2', '1× Refrigerator, 150W, 24 hours/day (Critical)'],
        ['Load 3', '1× Television (LED 42"), 80W, 5 hours/day (Important)'],
        ['Load 4', '2× Electric Fan, 60W each, 8 hours/day (Non-Critical)'],
        ['Battery Voltage', '24V'],
        ['PV Module', '400Wp | Voc = 49.5V | Isc = 10.36A | Vmp = 41.7V | Imp = 9.59A'],
        ['Peak Sun Hours', '5.5 hours/day (Southern Africa region)'],
        ['Autonomy Days', '2 days'],
        ['Depth of Discharge', '50% (lead-acid batteries)'],
        ['System Efficiency', '85%'],
        ['PV → Controller cable', '10 metres'],
        ['Controller → Battery cable', '2 metres'],
        ['Battery → Inverter cable', '1.5 metres'],
        ['Inverter → Load Panel cable', '5 metres'],
      ],
      [3500, 5526]
    ),
    spacer(200),
    h2('8.2 Step-by-Step Calculations'),
    h3('Daily Energy Consumption'),
    code('Lights:       4 × 10W × 6h = 240 Wh'),
    code('Refrigerator: 1 × 150W × 24h = 3600 Wh'),
    code('Television:   1 × 80W × 5h  = 400 Wh'),
    code('Fans:         2 × 60W × 8h  = 960 Wh'),
    code('─────────────────────────────────────'),
    code('TOTAL DAILY ENERGY = 5,200 Wh/day'),
    spacer(100),
    h3('PV Array Sizing'),
    code('Energy Required = 5200 / 0.85 = 6,118 Wh/day'),
    code('Panel Output    = 400 × 5.5 × 0.95 × 0.98 = 2,047 Wh/panel/day'),
    code('Raw Panels      = 6118 / 2047 = 2.99  → round up to 3'),
    code('Config: Panels in Series = CEIL(24 / 41.7) = 1   → use 2 for margin'),
    code('Config: Panels in Parallel = CEIL(3 / 2) = 2'),
    code('TOTAL PANELS = 2S × 2P = 4 panels   (1,600Wp array)'),
    spacer(100),
    h3('Battery Bank Sizing'),
    code('C_required = (5200 × 2) / (24 × 0.5 × 0.85)'),
    code('           = 10400 / 10.2 = 1,020 Ah'),
    code('Batteries in Series   = 24V / 12V = 2 (two 12V batteries in series)'),
    code('Batteries in Parallel = CEIL(1020 / 200) = 6 strings (200Ah each)'),
    code('TOTAL BATTERIES = 2S × 6P = 12 batteries'),
    code('TOTAL BANK CAPACITY = 6 × 200Ah = 1,200 Ah at 24V'),
    spacer(100),
    h3('Inverter Sizing'),
    code('Peak Load = (4×10) + 150 + 80 + (2×60) = 40+150+80+120 = 390W'),
    code('Surge     = 390 × 1.25 = 487.5W'),
    code('SELECTED INVERTER = 500W (next standard size above 487.5W)'),
    spacer(100),
    h3('Charge Controller'),
    code('Array Power = 4 × 400W = 1600W  (> 500W → MPPT selected)'),
    code('MPPT Rating = (1600 / 24) × 1.25 = 66.7A × 1.25 = 83.3A'),
    code('SELECTED CONTROLLER = 100A MPPT (next standard size above 83.3A)'),
    spacer(100),
    h3('PV to Controller Cable'),
    code('Current     = Isc × Parallel Strings × 1.25 = 10.36 × 2 × 1.25 = 25.9A'),
    code('A_min       = (0.0175 × 2×10 × 25.9) / (24 × 0.03)'),
    code('            = (0.0175 × 20 × 25.9) / 0.72'),
    code('            = 9.065 / 0.72 = 12.6 mm²'),
    code('SELECTED CABLE = 16 mm² (next standard above 12.6mm²)'),
    spacer(100),
    h3('Protection — PV Array Fuse'),
    code('Isc_array = Isc × Parallel = 10.36 × 2 = 20.72A'),
    code('Fuse      = 20.72 × 1.56 = 32.3A'),
    code('SELECTED FUSE = 40A (next standard rating above 32.3A)'),
    spacer(200),
    h2('8.3 Expected App Output Summary'),
    simpleTable(
      ['Component', 'Calculated Value', 'App Output'],
      [
        ['PV Panels', '4 panels (2S × 2P)', '4 panels — 2S × 2P'],
        ['Array Power', '1,600 Wp', '1600 W'],
        ['Battery Capacity', '1,200 Ah (12 batteries)', '1200 Ah — 12 batteries'],
        ['Battery Bank Energy', '28,800 Wh', '28,800 Wh'],
        ['Inverter Rating', '500W (Pure Sine Wave)', '500W'],
        ['Charge Controller', '100A MPPT', '100A MPPT'],
        ['PV Cable', '16 mm²', '16 mm²'],
        ['PV Fuse', '40A', '40A'],
      ],
      [2500, 2800, 3726]
    ),
    pageBreak(),
  ];

  // ====================== SECTION 9: TROUBLESHOOTING ======================
  const section9 = [
    h1('9. Troubleshooting Guide'),
    simpleTable(
      ['Error / Issue', 'Root Cause', 'Solution'],
      [
        ['Cannot find module node-fetch', 'Node.js v24 incompatible with Expo 50 analytics module', 'Project uses Expo 52. If issue persists, install Node.js v20 LTS from nodejs.org'],
        ['PluginError: expo-haptics invalid plugin', 'expo-haptics listed in plugins array in app.json', 'Remove expo-haptics from the plugins array in app.json (already fixed in this project)'],
        ['ERESOLVE dependency conflict', 'victory-native@40 requires React 19, incompatible with Expo', 'Project uses custom SVG charts. Do not add victory-native to package.json'],
        ['EPERM rmdir on Windows', 'Windows file system locks deeply-nested node_modules', 'Run in PowerShell: Remove-Item -Recurse -Force node_modules'],
        ['Metro bundler port 8081 in use', 'Another instance of Expo or Metro is running', 'Run: npx expo start --port 8082'],
        ['QR code not connecting', 'Phone and PC on different networks', 'Connect both devices to the same WiFi. Or use tunnel: npx expo start --tunnel'],
        ['App crashes on Android emulator', 'Emulator allocated insufficient RAM', 'In Android Studio AVD Manager, set RAM to 4096 MB'],
        ['EAS build fails: no android.package', 'app.json missing android package name', 'Ensure "package": "com.solarpv.pro" is in the android section of app.json'],
        ['Gradle build fails: SDK not found', 'ANDROID_HOME not set or Android SDK not installed', 'Open Android Studio → SDK Manager and install Android SDK 34. Set ANDROID_HOME in Windows Environment Variables'],
        ['App shows white screen on launch', 'JavaScript bundle error during startup', 'Run: npx expo start, check red error overlay for the specific error message'],
        ['Calculations show NaN values', 'One or more input fields left blank or set to zero', 'All numeric inputs must be greater than zero before calculating'],
        ['npm install is very slow', 'Network throttling or proxy interference', 'Try: npm install --legacy-peer-deps --prefer-offline after first install'],
      ],
      [2800, 2800, 3426]
    ),
    pageBreak(),
  ];

  // ====================== SECTION 10: STANDARDS ======================
  const section10 = [
    h1('10. Engineering Standards Reference'),
    para('The SolarPV Pro application implements calculations and safety factors based on the following internationally recognised electrical and solar PV standards:'),
    spacer(100),
    simpleTable(
      ['Standard', 'Title', 'Application in App'],
      [
        ['IEC 62548:2016', 'PV Arrays – Design Requirements', 'Array sizing, series/parallel configuration, protection requirements'],
        ['IEC 60364-7-712', 'Low-Voltage Installations – PV Supply Systems', 'Cable sizing, voltage drop limits (3% DC), earthing'],
        ['NEC Article 690', 'Solar Photovoltaic (PV) Systems', 'PV fuse sizing (1.56 × Isc), wire ampacity, rapid shutdown'],
        ['BS 7671:2018', 'Requirements for Electrical Installations (18th Edition)', 'Cable current ratings, protection coordination'],
        ['IEC 61427', 'Secondary Cells for PV Energy Systems', 'Battery DoD limits, charging efficiency factors'],
        ['IEEE 1547', 'Interconnection of Distributed Energy Resources', 'Inverter specifications, anti-islanding reference'],
        ['IEC 60269', 'Low-Voltage Fuses', 'Standard fuse ratings selection, breaking capacity'],
        ['IEC 61730', 'PV Module Safety Qualification', 'Temperature derating coefficients, safety classifications'],
      ],
      [2200, 3000, 3826]
    ),
    pageBreak(),
  ];

  // ====================== APPENDICES ======================
  const appendices = [
    h1('Appendix A: Quick Engineering Reference Card'),
    para('Use this table for rapid manual verification of app results:'),
    spacer(100),
    simpleTable(
      ['Parameter', 'Formula', 'Safety Factor', 'Rounding'],
      [
        ['PV Panels', 'E_required / (PSH × P_module × η_soiling × η_mismatch)', 'None', 'Round UP'],
        ['Battery (Ah)', '(E_daily × Days) / (V_bat × DoD × η_battery)', 'None', 'Round UP'],
        ['Inverter (W)', 'Σ(Load Power) × Surge Factor', '1.25×', 'Next standard size'],
        ['Controller (A) MPPT', '(Array Power / V_battery) × 1.25', '1.25×', 'Next standard size'],
        ['Controller (A) PWM', 'Isc_corrected × N_parallel × 1.25', '1.25×', 'Next standard size'],
        ['Cable Size (mm²)', '(ρ × 2L × I) / (V × Δv_max)', 'Built into current', 'Next standard size'],
        ['PV Array Fuse (A)', 'Isc_array × 1.56', '1.56× per NEC 690.9', 'Next standard rating'],
        ['Battery Breaker (A)', '(P_inv / (V_bat × 0.90)) × 1.25', '1.25×', 'Next standard rating'],
        ['AC Breaker (A)', '(P_inv / 230V) × 1.25', '1.25×', 'Next standard rating'],
      ],
      [2000, 3200, 2000, 1826]
    ),
    spacer(400),
    h1('Appendix B: Regional Peak Sun Hours'),
    para('Use these values for quick entry in Design Wizard Step 3. Values represent average daily peak sun hours (PSH) across each region.'),
    spacer(100),
    simpleTable(
      ['Region', 'Peak Sun Hours', 'Notes'],
      [
        ['Southern Africa', '5.5 hours', 'South Africa, Zimbabwe, Zambia, Mozambique'],
        ['East Africa', '5.0 hours', 'Kenya, Tanzania, Uganda, Ethiopia'],
        ['West Africa', '5.2 hours', 'Ghana, Nigeria, Senegal, Ivory Coast'],
        ['North Africa', '6.0 hours', 'Egypt, Libya, Morocco, Algeria'],
        ['Middle East', '6.5 hours', 'Saudi Arabia, UAE, Jordan, Iraq'],
        ['South Asia', '5.0 hours', 'India, Bangladesh, Sri Lanka, Pakistan'],
        ['Southeast Asia', '4.5 hours', 'Indonesia, Philippines, Vietnam, Malaysia'],
        ['Southern Europe', '5.0 hours', 'Spain, Italy, Greece, Portugal'],
        ['Northern Europe', '3.0 hours', 'UK, Germany, France, Netherlands'],
        ['North America (South)', '5.5 hours', 'Texas, California, Florida, Arizona'],
        ['North America (North)', '4.0 hours', 'Canada, Northern US states'],
        ['South America', '5.0 hours', 'Brazil, Colombia, Peru, Chile'],
        ['Australia', '5.5 hours', 'Australia-wide average'],
      ],
      [2800, 2000, 4226]
    ),
    spacer(400),
    h1('Appendix C: Standard Component Sizes Reference'),
    h2('Cable Sizes (IEC 60228)'),
    para('Standard copper cable cross-sections available for selection:'),
    code('1.5 | 2.5 | 4 | 6 | 10 | 16 | 25 | 35 | 50 | 70 | 95 | 120 mm²'),
    spacer(100),
    h2('Breaker / Fuse Ratings'),
    code('6 | 10 | 16 | 20 | 25 | 32 | 40 | 50 | 63 | 80 | 100 | 125 | 160 | 200 A'),
    spacer(100),
    h2('Standard Inverter Sizes'),
    code('300 | 500 | 800 | 1000 | 1500 | 2000 | 3000 | 5000 | 8000 | 10000 W'),
    spacer(100),
    h2('Standard Charge Controller Sizes'),
    code('10 | 20 | 30 | 40 | 50 | 60 | 80 | 100 | 150 | 200 A'),
  ];

  return [
    ...coverPage,
    ...tocSection,
    ...section1,
    ...section2,
    ...section3,
    ...section4,
    ...section5,
    ...section6,
    ...section7,
    ...section8,
    ...section9,
    ...section10,
    ...appendices,
  ];
}

// ============================================================
// BUILD AND WRITE
// ============================================================

const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '\u2022',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 560, hanging: 280 } } },
        }, {
          level: 1,
          format: LevelFormat.BULLET,
          text: '\u25E6',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 900, hanging: 280 } } },
        }],
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0,
          format: LevelFormat.DECIMAL,
          text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 560, hanging: 280 } } },
        }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22, color: DARK_TEXT } },
    },
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: TEAL },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: NAVY },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 },
      },
      {
        id: 'Heading3',
        name: 'Heading 3',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: TEAL },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
              children: [
                new TextRun({ text: 'SolarPV Pro — Technical Guide & User Manual', font: 'Arial', size: 18, color: '666666' }),
                new TextRun({ text: '\t', font: 'Arial', size: 18 }),
                new TextRun({ text: 'EEE3206  |  Department of Electronic Engineering', font: 'Arial', size: 18, color: TEAL }),
              ],
              tabStops: [{ type: 'right', position: CONTENT_W }],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
              children: [
                new TextRun({ text: 'Version 1.0.0  |  April 2026', font: 'Arial', size: 16, color: '999999' }),
                new TextRun({ text: '\t', font: 'Arial', size: 16 }),
                new TextRun({ text: 'Page ', font: 'Arial', size: 16, color: '666666' }),
                new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: TEAL, bold: true }),
              ],
              tabStops: [{ type: 'right', position: CONTENT_W }],
            }),
          ],
        }),
      },
      children: buildDoc(),
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const outputPath = 'C:\\Users\\TechCharities\\Documents\\final_project\\SolarPVApp_Guide.docx';
  fs.writeFileSync(outputPath, buffer);
  console.log('✅ Document created successfully!');
  console.log('📄 Saved to:', outputPath);
  const stats = fs.statSync(outputPath);
  console.log('📦 File size:', (stats.size / 1024).toFixed(1), 'KB');
}).catch(err => {
  console.error('❌ Error:', err.message);
  process.exit(1);
});
