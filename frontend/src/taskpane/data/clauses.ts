export interface Clause {
  id: string;
  label: string;
  text: string;
  isMandatory?: boolean;
}

export interface ReportSection {
  [key: string]: Clause[];
}

export const reportLibrary: { [key: string]: ReportSection } = {
  general_legal: {
    tenure: [
      {
        id: "ten_free",
        label: "Freehold/Vacant",
        text: "It has been assumed that the property is being sold on a Freehold/Leasehold basis with vacant possession on completion of sale. This should be confirmed by your legal advisor.",
        isMandatory: true,
      },
    ],
    sewer_party: [
      {
        id: "leg_septic_2020",
        label: "Septic Tank (Jan 2020)",
        text: "Since 1st January 2020 all septic tanks need to be replaced with a sewerage treatment plant with compliant water run off discharge into the surrounding ground, this is normally by a network of perforated below ground pipes.",
      },
      {
        id: "pw_act_1996",
        label: "Party Wall Act 1996",
        text: "Since 1st July 1997, this Act has obliged anyone undertaking works of a structural nature to or near a shared boundary to notify all adjoining owners. Such works include the installation of beams, damp proofing courses, and excavating.",
      },
    ],
  },
  exterior_structure: {
    walls_roof: [
      {
        id: "wall_clay",
        label: "Clay Subsoil Alert",
        text: "Your attention is drawn to the fact that the subsoil in this district is predominantly clay. Clay subsoils are susceptible to shrinkage during periods of extremely dry weather. Roots from trees and shrubs can also have a significant contributory effect.",
      },
      {
        id: "rf_pitched",
        label: "Pitched Slate/Tile",
        text: "The property has been constructed with a pitched roof structure covered with natural slate or concrete tiles.",
      },
    ],
    hazards: [
      {
        id: "saf_lead_paint",
        label: "Lead Paint (Pre-1992)",
        text: "Lead paint was phased out in the UK in the 1960s but was not fully banned until 1992, so any house that predates this could possibly have lead in the painted surfaces. Suitable precautions should be taken when rubbing down.",
      },
      {
        id: "bound_danger",
        label: "Dangerous Boundary Wall",
        text: "The boundary walls are approximately 1m/2m high but only 112mm in thickness. Freestanding garden walls of this slenderness should not exceed a height of 900mm. It is strongly recommended that the wall be taken down and rebuilt as a matter of urgency.",
      },
    ],
  },
  services_utilities: {
    mech_elec: [
      {
        id: "wat_lead_main",
        label: "Lead Water Main",
        text: "The incoming main is in lead, a material which is hazardous to health. It is recommended that the existing main be stripped out and a new individual main installed in blue polyethylene.",
      },
      {
        id: "elec_rccb",
        label: "RCCB Protection",
        text: "The installation is fitted with a Residual Current Circuit Breaker. This is a modern system designed to protect the users from electric shock. RCCB’s are extremely sensitive and consequently occasional tripping will occur.",
      },
    ],
    heating: [
      {
        id: "heat_combo",
        label: "Combo Boiler Logic",
        text: "This is a modern combination boiler appliance which removes the need for cold and hot water storage. Output can vary with changes in pressure; filling of baths can take a long time in comparison with conventional systems.",
      },
    ],
  },
  appendix_guidance: {
    essential: [
      {
        id: "gui_testing",
        label: "Technical Hotlines",
        text: "ESSENTIAL GUIDANCE:\n• Electrical Systems: N.I.C.E.I.C 020 7564 2323.\n• Gas Appliances: 'Gas Safe' registered specialist 01256 372200.\n• Radon: Health Protection Agency www.hpa.org.uk.",
      },
    ],
    terms: [
      {
        id: "tc_scope",
        label: "Engagement Terms (5.4)",
        text: "TERMS AND CONDITIONS:\n5.4. Generally: The Surveyor will inspect diligently but is not required to undertake any action which would risk damage to the Property.\n11.4. Dispute Resolution: In the event that the Client has a complaint, a formal complaints handling procedure will be followed.",
        isMandatory: true,
      },
    ],
  },
};
