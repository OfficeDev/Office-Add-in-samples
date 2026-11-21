// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

export const members = [
  { name: "Consultant A", rate: 150 },
  { name: "Consultant B", rate: 130 },
  { name: "Consultant C", rate: 120 },
  { name: "Consultant D", rate: 120 },
  { name: "Consultant E", rate: 180 },
  { name: "Consultant F", rate: 100 },
];

export interface TeamMember {
  name: string;
  rate: number;
}

export const roles = ["Lead", "UI designer", "Web engineer", "Database engineer"];

export function calculateTeam(team: readonly TeamMember[], hours: number) {
  if (team.length !== roles.length) {
    throw new Error("Select a consultant for each of the four roles.");
  }
  if (!Number.isFinite(hours) || hours <= 0 || hours > 10000) {
    throw new Error("Hours per person must be greater than 0 and no more than 10,000.");
  }
  if (
    team.some((member) => !member.name.trim() || !Number.isFinite(member.rate) || member.rate < 0)
  ) {
    throw new Error("Each consultant must have a name and a valid nonnegative rate.");
  }
  const combinedRate = team.reduce((sum, member) => sum + member.rate, 0);
  return { averageRate: combinedRate / team.length, total: combinedRate * hours };
}

export function currency(value: number): string {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value);
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function cell(value: string): string {
  return `<w:tc><w:p><w:r><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r></w:p></w:tc>`;
}

export function createTeamOoxml(team: readonly TeamMember[], hours: number): string {
  const { averageRate, total } = calculateTeam(team, hours);
  const rows = [
    ["Role", "Consultant", "Hourly rate (USD)"],
    ...team.map((member, index) => [roles[index], member.name, currency(member.rate)]),
    ["Average hourly rate", "", currency(averageRate)],
    ["Hours per person", "", String(hours)],
    ["Total (USD)", "", currency(total)],
  ];
  return `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Project team</w:t></w:r></w:p>
      <w:tbl><w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
        <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
        <w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
      </w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/><w:gridCol w:w="2400"/></w:tblGrid>
        ${rows.map((row) => `<w:tr>${row.map(cell).join("")}</w:tr>`).join("\n")}
      </w:tbl><w:p/>
    </w:body></w:document></pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}
