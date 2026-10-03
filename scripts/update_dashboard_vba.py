import re

with open('vba/modDashboardUIUX.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# Replace BuildChartContainer calls for CC
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CC_HourlyVolume"[^"]+"Column Chart")',
    r'\1, "icon_hourly_surge.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CC_TopicBreakdown"[^"]+"Clustered Bar")',
    r'\1, "icon_topic_sla.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CC_AgentQuadrant"[^"]+"Scatter Plot")',
    r'\1, "icon_agent_quadrant.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CC_AgentScorecard"[^"]+"Matrix Table")',
    r'\1, "icon_audit_matrix.svg"',
    text
)

# Churn containers
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CH_ContractRisk"[^"]+"Column Chart")',
    r'\1, "icon_retention_users.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CH_TenureCohort"[^"]+"Area / Line Chart")',
    r'\1, "icon_hourly_surge.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CH_PaymentFriction"[^"]+"Clustered Bar")',
    r'\1, "icon_topic_sla.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"CH_ServiceMatrix"[^"]+"Matrix Table")',
    r'\1, "icon_audit_matrix.svg"',
    text
)

# Diversity containers
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"DI_PipelineFunnel"[^"]+"Funnel / Bar")',
    r'\1, "icon_diversity_parity.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"DI_DeptParity"[^"]+"Clustered Column")',
    r'\1, "icon_topic_sla.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"DI_PromoVelocity"[^"]+"Bar Chart")',
    r'\1, "icon_agent_quadrant.svg"',
    text
)
text = re.sub(
    r'(BuildChartContainer\s+ws,\s+"DI_PerformanceAudit"[^"]+"Matrix Table")',
    r'\1, "icon_audit_matrix.svg"',
    text
)

with open('vba/modDashboardUIUX.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print("All visual container calls successfully upgraded with SVG vector icons!")
