export type Measure = {
    _id: string,
    artefacts: Artefact[],
    risks: Risk[]
    monthlySpendings: number[],
    kpiData: KPIData,
    id: number,
    title: string,
    name: string,
    description: string,
    time: string,
    lastUpdate: string,
    focusArea: string,
    focusAreaFull: string,
    measureLead: string,
    measureSponsor: string,
    lineOrgSponsor: string,
    solutionManager: string,
    approved: number,
    spent: number,
    kpiName: string,
    actuals: number,
    target: number,
    totalApprovedBudget: number,
    artefact: number,
    budgetDetail: BudgetDetail
}

export type BudgetDetail = {
    totalApprovedBudget: number,
    spentBudget: number,
    invoicedBudget: number,
    forecastBudget: number,
    contractBudget: number
}


export type KPIData = {
    target: number,
    actuals: number,
    baseline: number,
    plan1: number,
    plan2: number,
    plan3: number,
    plan4: number,
}

export type Risk = {
    risk: string | number,
    description: string | number,
    criticality: string | number,
    migration: string | number,
    resolutionDate: string | number,
}

export type Artefact = {
    _id: string,
    id: number,
    description: string,
    progress: number,
    budget: string,
    achievement: string,
    work: string,
}

export type Overview = {
    _id: string,
    name: string,
    measures: string[],
    kpiProgress: number,
    overallStatus: number,
    progress: number,
    totalBudget: number,
}

export type Budget = {
    _id: string,
    monthlySpendings: number[],
    approvedBudgetPerMonth: number,
    year: number,
}

export type InitialOverview = {
    numberOfMeasures: number,
    totalBudget: number,
    overallStatus: number,
    progressOverviewBarResult: number,
    KPIProgressResult: number,
}

export type ParseOverview = {
    row: number,
    name: string,
    risk: number,
    budget: number,
    artefact: number
}

export type KpiProgressData = {
    measureName: string,
    progress: number
}

export type AllBudgetMeasures = { [x: number]: number }

export type KPI = {
    target: number,
    actuals: number,
    baseline: number,
    plan1: number,
    plan2: number,
    plan3: number,
    plan4: number,
}

export type SheetType = {
    [key: string]: string | number
}