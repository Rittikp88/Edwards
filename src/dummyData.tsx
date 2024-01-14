import { Data } from "./Home";

export const dummyData: Data = {
  materials: [
    {
      id: "m1",
      name: "Aluminum Alloy",
      type: "material",
      studies: ["s1", "s2", "s3", "s4", "s5", "s6"],
    },
    {
      id: "m2",
      name: "Carbon Fiber",
      type: "material",
      studies: ["s3"],
    },
    {
      id: "m3",
      name: "Titanium",
      type: "material",
      studies: ["s5", "s6", "s7", "s8", "s9", "s10"],
    },
    {
      id: "m4",
      name: "Stainless Steel",
      type: "material",
      studies: ["s7", "s8", "s9", "s10", "s11", "s12"],
    },
    {
      id: "m5",
      name: "Polymer Composite",
      type: "material",
      studies: ["s9", "s10", "s11", "s12", "s13", "s14"],
    },
    {
      id: "m6",
      name: "Graphene",
      type: "material",
      studies: ["s11", "s12", "s13", "s14"],
    },
    {
      id: "m7",
      name: "Ceramic",
      type: "material",
      studies: ["s13", "s14"],
    },
  ],
  studies: [
    {
      id: "s1",
      name: "Durability Analysis",
      type: "study",
      protocols: ["p1", "p2", "p3", "p4", "p5", "p6"],
    },
    {
      id: "s2",
      name: "Thermal Resistance Study",
      type: "study",
      protocols: ["p3", "p4", "p5", "p6", "p7"],
    },
    {
      id: "s3",
      name: "Aerodynamic Efficiency",
      type: "study",
      protocols: ["p5"],
    },
    {
      id: "s4",
      name: "Material Fatigue Assessment",
      type: "study",
      protocols: ["p7", "p8", "p9", "p10", "p11", "p12"],
    },
    {
      id: "s5",
      name: "Impact Resistance",
      type: "study",
      protocols: ["p9", "p10", "p11", "p12", "p13", "p14"],
    },
    {
      id: "s6",
      name: "Flexibility Study",
      type: "study",
      protocols: ["p11", "p12", "p13", "p14", "p15", "p16"],
    },
    {
      id: "s7",
      name: "Corrosion Study",
      type: "study",
      protocols: ["p13", "p14", "p15", "p16", "p17", "p18"],
    },
    {
      id: "s8",
      name: "UV Resistance Analysis",
      type: "study",
      protocols: ["p15", "p16", "p17", "p18", "p19", "p20"],
    },
    {
      id: "s9",
      name: "Electrical Conductivity Study",
      type: "study",
      protocols: ["p17", "p18", "p19", "p20"],
    },
    {
      id: "s10",
      name: "Acoustic Damping Study",
      type: "study",
      protocols: ["p19", "p20"],
    },
    {
      id: "s11",
      name: "High Temperature Resistance Study",
      type: "study",
      protocols: ["p19", "p20"],
    },
    {
      id: "s12",
      name: "Low Temperature Performance Assessment",
      type: "study",
      protocols: ["p19", "p20"],
    },
    {
      id: "s13",
      name: "Radiation Shielding Effectiveness",
      type: "study",
      protocols: ["p19", "p20"],
    },
    {
      id: "s14",
      name: "Microstructural Analysis",
      type: "study",
      protocols: ["p19", "p20"],
    },
  ],
  protocols: [
    {
      id: "p1",
      name: "Tensile Strength Test",
      type: "protocol",
      tests: ["t1", "t2"],
    },
    {
      id: "p2",
      name: "Compression Test",
      type: "protocol",
      tests: ["t3", "t4"],
    },
    {
      id: "p3",
      name: "Thermal Cycling",
      type: "protocol",
      tests: ["t5", "t6"],
    },
    {
      id: "p4",
      name: "Corrosion Resistance Test",
      type: "protocol",
      tests: ["t7", "t8"],
    },
    {
      id: "p5",
      name: "Wind Tunnel Testing",
      type: "protocol",
      tests: ["t9"],
    },
    {
      id: "p6",
      name: "Vibration Analysis",
      type: "protocol",
      tests: ["t11", "t12"],
    },
    {
      id: "p7",
      name: "Fatigue Testing",
      type: "protocol",
      tests: ["t13", "t14"],
    },
    {
      id: "p8",
      name: "UV Exposure Test",
      type: "protocol",
      tests: ["t15", "t16"],
    },
    { id: "p9", name: "Drop Test", type: "protocol", tests: ["t17", "t18"] },
    {
      id: "p10",
      name: "Bending Test",
      type: "protocol",
      tests: ["t19", "t20"],
    },
    {
      id: "p11",
      name: "Electrochemical Test",
      type: "protocol",
      tests: ["t21", "t22"],
    },
    {
      id: "p12",
      name: "Thermal Conductivity Test",
      type: "protocol",
      tests: ["t23", "t24"],
    },
    {
      id: "p13",
      name: "Electromagnetic Interference Test",
      type: "protocol",
      tests: ["t25", "t26"],
    },
    {
      id: "p14",
      name: "Acoustic Emission Test",
      type: "protocol",
      tests: ["t27", "t28"],
    },
    {
      id: "p15",
      name: "UV Accelerated Aging Test",
      type: "protocol",
      tests: ["t29", "t30"],
    },
    {
      id: "p16",
      name: "Salt Fog Test",
      type: "protocol",
      tests: ["t31", "t32"],
    },
    {
      id: "p17",
      name: "X-ray Diffraction Analysis",
      type: "protocol",
      tests: ["t33", "t34"],
    },
    {
      id: "p18",
      name: "Scanning Electron Microscopy",
      type: "protocol",
      tests: ["t35", "t36"],
    },
    {
      id: "p19",
      name: "Thermogravimetric Analysis",
      type: "protocol",
      tests: ["t37", "t38"],
    },
    {
      id: "p20",
      name: "Dynamic Mechanical Analysis",
      type: "protocol",
      tests: ["t39", "t40"],
    },
  ],
  tests: [
    {
      id: "t1",
      name: "Tensile Test - Low Load",
      type: "test",
      duration: 30,
    },
    {
      id: "t2",
      name: "Tensile Test - High Load",
      type: "test",
      duration: 45,
    },
    {
      id: "t3",
      name: "Compression Test - Low Pressure",
      type: "test",
      duration: 35,
    },
    {
      id: "t4",
      name: "Compression Test - High Pressure",
      type: "test",
      duration: 50,
    },
    {
      id: "t5",
      name: "Thermal Shock Test",
      type: "test",
      duration: 60,
    },
    {
      id: "t6",
      name: "Thermal Cycle Test",
      type: "test",
      duration: 75,
    },
    {
      id: "t7",
      name: "Salt Spray Test",
      type: "test",
      duration: 120,
    },
    {
      id: "t8",
      name: "Humidity Test",
      type: "test",
      duration: 90,
    },
    {
      id: "t9",
      name: "Low-speed Wind Tunnel Test",
      type: "test",
      duration: 180,
    },
    {
      id: "t10",
      name: "High-speed Wind Tunnel Test",
      type: "test",
      duration: 240,
    },
    {
      id: "t11",
      name: "Random Vibration Test",
      type: "test",
      duration: 150,
    },
    {
      id: "t12",
      name: "Harmonic Vibration Test",
      type: "test",
      duration: 110,
    },
    {
      id: "t13",
      name: "Rotating Bend Fatigue Test",
      type: "test",
      duration: 200,
    },
    {
      id: "t14",
      name: "Axial Fatigue Test",
      type: "test",
      duration: 170,
    },
    {
      id: "t15",
      name: "UV Aging Test",
      type: "test",
      duration: 130,
    },
    {
      id: "t16",
      name: "Accelerated Weathering Test",
      type: "test",
      duration: 160,
    },
    { id: "t17", name: "Ball Drop Test", type: "test", duration: 20 },
    { id: "t18", name: "Edge Drop Test", type: "test", duration: 25 },
    { id: "t19", name: "Three-Point Bending Test", type: "test", duration: 40 },
    { id: "t20", name: "Four-Point Bending Test", type: "test", duration: 35 },
    {
      id: "t21",
      name: "Potentiodynamic Polarization Test",
      type: "test",
      duration: 50,
    },
    { id: "t22", name: "Galvanostatic Test", type: "test", duration: 45 },
    { id: "t23", name: "Laser Flash Analysis", type: "test", duration: 55 },
    { id: "t24", name: "Heat Flow Meter Test", type: "test", duration: 60 },
    { id: "t25", name: "Radiated Emissions Test", type: "test", duration: 70 },
    { id: "t26", name: "Conducted Emissions Test", type: "test", duration: 65 },
    { id: "t27", name: "Linear Location Test", type: "test", duration: 75 },
    {
      id: "t28",
      name: "Zonal Discrimination Test",
      type: "test",
      duration: 80,
    },
    {
      id: "t29",
      name: "UV Accelerated Aging - Low Intensity",
      type: "test",
      duration: 100,
    },
    {
      id: "t30",
      name: "UV Accelerated Aging - High Intensity",
      type: "test",
      duration: 150,
    },
    {
      id: "t31",
      name: "Salt Fog Exposure - Short Term",
      type: "test",
      duration: 80,
    },
    {
      id: "t32",
      name: "Salt Fog Exposure - Long Term",
      type: "test",
      duration: 200,
    },
    {
      id: "t33",
      name: "X-ray Crystallography",
      type: "test",
      duration: 60,
    },
    {
      id: "t34",
      name: "X-ray Fluorescence Analysis",
      type: "test",
      duration: 70,
    },
    {
      id: "t35",
      name: "SEM Surface Analysis",
      type: "test",
      duration: 90,
    },
    {
      id: "t36",
      name: "SEM Internal Structure Analysis",
      type: "test",
      duration: 110,
    },
    {
      id: "t37",
      name: "Thermogravimetric Degradation Test",
      type: "test",
      duration: 120,
    },
    {
      id: "t38",
      name: "Thermal Stability Analysis",
      type: "test",
      duration: 140,
    },
    {
      id: "t39",
      name: "DMA Stress-Strain Test",
      type: "test",
      duration: 160,
    },
    {
      id: "t40",
      name: "DMA Temperature Sweep",
      type: "test",
      duration: 180,
    },
  ],
  resources:  [
    {
      id: "r1",
      name: "Alex",
      status: "Free"
    },
    {
      id: "r1",
      name: "Jonny",
      status: "Free"
    },
    {
      id: "r1",
      name: "Ester",
      status: "Free"
    },
    {
      id: "r1",
      name: "Heena",
      status: "Free"
    },
    {
      id: "r1",
      name: "Molex",
      status: "Free"
    },
    {
      id: "r1",
      name: "Nirobi",
      status: "Free"
    },
    {
      id: "r1",
      name: "Jaspin",
      status: "Free"
    },
    {
      id: "r1",
      name: "Jessica",
      status: "Free"
    },
    {
      id: "r1",
      name: "Alex",
      status: "Free"
    },


 ]

  
};
