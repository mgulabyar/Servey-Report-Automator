export interface ReportSection {
  id: string;
  title: string;
  content: string;
  isMandatory: boolean;
  category: string;
}

export const sectionsData: ReportSection[] = [
  {
    id: "header_intro",
    title: "Building Survey Report Intro",
    category: "Introduction",
    isMandatory: true,
    content: "BUILDING SURVEY REPORT\nMaking the most of your report...\n[Full Intro Text Here]",
  },
  {
    id: "condition_ratings",
    title: "Condition Ratings 1, 2 & 3",
    category: "Guidance",
    isMandatory: false,
    content:
      "CONDITION RATINGS 1,2 & 3\nWhat everyone wants to know is how significant any defect is...",
  },
  {
    id: "summary_1_1",
    title: "1.1 Property Details",
    category: "Summary",
    isMandatory: true,
    content:
      "1.0 SUMMARY\n1.1 PROPERTY\nProperty address: [Address]\nProperty type: [Description]...",
  },
  {
    id: "mining_info",
    title: "Coal Mining Area Info",
    category: "Environmental",
    isMandatory: false,
    content:
      "The property is situated on the outskirts of an historic coal mining area. The purchasing solicitor should obtain a mining search on the property.",
  },
];
