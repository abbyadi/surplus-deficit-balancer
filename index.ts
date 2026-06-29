Office.onReady(() => {
  document.getElementById("run")!.addEventListener("click", () => tryCatch(createObjFnx, balanceUses));
  document.getElementById("soft-debt")!.addEventListener("click", () => tryCatch(createObjFnx, balanceSources));
  document.getElementById("Max-Fed-Credits")!.addEventListener("click", () => tryCatch(createObjFnx, maxFedCredits));
  document.getElementById("Fed-Credit-Adjuster")!.addEventListener("click", () => tryCatch(createObjFnx, fedCreditAdjstr));
  document.getElementById("State-Credit-Adjuster")!.addEventListener("click", () => tryCatch(createObjFnx, stateCredAdjstr));
  document.getElementById("defer-fee")!.addEventListener("click", () => tryCatch(createObjFnx, deferFee));
});

function modifyStr(str: string): string {
  let strLowerCase = str.toLowerCase();
  strLowerCase = strLowerCase.replace(/[^a-zA-Z0-9]+/g, ".*");
  return strLowerCase;
}

function getInputValue(id: string): string {
  return (document.getElementById(id) as HTMLInputElement).value;
}

async function createObjFnx(fnx: (obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) => Promise<void>) {
  await Excel.run(async (context) => {
    const ss = context.workbook;
    const namedRange = getInputValue("range-name");
    const rangeNames: Record<string, string> = {
      is4Pct: "Is4PctCredits",
      isAcqRehab: "Is_AcqRehab",
      surplusShortfall: "SurplusShortfall",
      rngDevCosts: namedRange,
      rngSoftDebt: "SoftDebt",
      volReductionState: "VolCreditReduction_State",
      volReductionFed: "VolCreditReduction_Fed",
      currStateCreditAmt: "StateCreditAmt",
      eligibleBasis: "EligibleBasis",
      lpEquity: "LPEquity",
      lpProfitLoss: "LP_ProfitLoss",
      fedPricing: "FedPricing",
      applicableFrc: "Applicable_Fraction",
      statePricing: "StatePricing",
      stateCreditRate: "StateCreditRate",
      ddaQctStatus: "DDAQCT_Status",
      cashFeeLimit: "Owner_CashFeeLimit",
      maxDevFee: "MaxDevFee",
    };

    const obj: Record<string, Excel.Range> = Object.fromEntries(
      Object.entries(rangeNames).map(([key, name]) => [
        key,
        ss.names.getItem(name).getRange(),
      ])
    );

    const loadConfig: Record<string, string[]> = {
      default: ["values", "address", "formulas", "formulasR1C1"],
      rngDevCosts: ["values", "rowIndex", "columnIndex"],
      rngSoftDebt: ["values", "rowIndex", "columnIndex"],
      volReductionState: ["values", "formulas"],
      volReductionFed: ["values", "formulas", "address"],
      eligibleBasis: ["values", "address"],
      cashFeeLimit: ["values", "address", "formulas", "formulasR1C1"],
    };

    Object.entries(obj).forEach(([key, range]) => {
      range.load(loadConfig[key] || loadConfig["default"]);
    });

    await context.sync();
    await fnx(obj, context);
  });
}

async function balanceUses(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const sheet = ctx.workbook.worksheets.getItem("Assumptions");
  const lineItemName = getInputValue("line-item-name").toLowerCase();
  const regExEscalation = new RegExp(modifyStr(lineItemName), "i");
  const isRehab = obj.isAcqRehab.values[0][0];
  let newRange: Excel.Range | undefined;
  let newRangePctFixed: Excel.Range | undefined;
  let newRangePctCalc: Excel.Range | undefined;
  const lineItemToChk = [
    "contingency escalation",
    "contingency design estimating",
    "contingency owners construction",
    "soft cost contingency",
    "hard costs unit construction",
  ];
  let counter = 0;

  for (let i = 0; i < obj.rngDevCosts.values.length; i++) {
    if (regExEscalation.test(obj.rngDevCosts.values[i][0].replace(/[^a-zA-Z0-9]+/g, ""))) {
      if (lineItemToChk.indexOf(lineItemName) !== -1) {
        if (isRehab === "Y" && counter === 0 && lineItemName !== lineItemToChk[3]) {
          counter++;
          continue;
        } else {
          if (lineItemToChk.indexOf(lineItemName) === 3) {
            newRange = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 6, 1, 1);
            newRangePctFixed = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 3, 1, 1);
            newRangePctCalc = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 4, 1, 1);
            newRange.load("values, address, formulas");
            newRangePctFixed.load("values, formulas");
            newRangePctCalc.load("values, address, formulas, formulasR1C1");
          } else if (lineItemToChk.indexOf(lineItemName) === 4) {
            newRange = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 1, 1, 1);
            newRange.load("values, address, formulas");
          } else {
            newRange = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 5, 1, 1);
            newRangePctFixed = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 3, 1, 1);
            newRangePctCalc = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 4, 1, 1);
            newRange.load("values, address, formulas");
            newRangePctFixed.load("values, formulas");
            newRangePctCalc.load("values, address, formulas, formulasR1C1");
          }
        }
      } else {
        newRange = sheet.getRangeByIndexes(obj.rngDevCosts.rowIndex + i, obj.rngDevCosts.columnIndex + 1, 1, 1);
        newRange.load("values, address, formulas");
      }
      break;
    }
  }

  if (!newRange) {
    console.error(`No matching line item found for: ${lineItemName}`);
    return;
  }

  await ctx.sync();

  const valSurplusShortfall = obj.surplusShortfall.values[0][0];
  let valEscalation = newRange.values[0][0];

  if (lineItemToChk.indexOf(lineItemName) !== -1) {
    if (valSurplusShortfall !== 0) {
      valEscalation = valEscalation + valSurplusShortfall;
      newRange.formulas = [[`=max(round(${valEscalation},0),0)`]];
      if (newRangePctFixed) {
        newRangePctFixed.values = [["Fixed"]];
      }
      if (lineItemToChk.indexOf(lineItemName) === 0) {
        if (isRehab === "N") {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-1]C[11]:R[-20]C[11],"Y",R[-1]C[14]:R[-20]C[14])+if(R[-53]C[11]="Y",R[-53]C[14]))`]];
        } else {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-1]C[11]:R[-20]C[11],"Y",R[-1]C[14]:R[-20]C[14])+if(R[-76]C[11]="Y",R[-77]C[14]))`]];
        }
      } else if (lineItemToChk.indexOf(lineItemName) === 1) {
        if (isRehab === "N") {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-2]C[12]:R[-21]C[12],"Y",R[-2]C[14]:R[-21]C[14])+if(R[-54]C[12]="Y",R[-54]C[14]))`]];
        } else {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-2]C[12]:R[-21]C[12],"Y",R[-2]C[14]:R[-21]C[14])+if(R[-77]C[12]="Y",R[-78]C[14]))`]];
        }
      } else if (lineItemToChk.indexOf(lineItemName) === 2) {
        if (isRehab === "N") {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-1]C[13]:R[-22]C[13],"Y",R[-1]C[14]:R[-22]C[14])+if(R[-55]C[13]="Y",R[-55]C[14]))`]];
        } else {
          newRangePctCalc!.formulasR1C1 = [[`=RC[1]/(sumif(R[-1]C[13]:R[-22]C[13],"Y",R[-1]C[14]:R[-22]C[14])+if(R[-78]C[13]="Y",R[-79]C[14]))`]];
        }
      } else if (lineItemToChk.indexOf(lineItemName) === 3) {
        newRangePctCalc!.formulasR1C1 = [[`=RC[2]/RC[1]`]];
      }
    }
  } else {
    if (valSurplusShortfall !== 0) {
      valEscalation = valEscalation + valSurplusShortfall;
      newRange.formulas = [[`=max(round(${valEscalation},0),0)`]];
    }
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function balanceSources(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const sheetAssumptions = ctx.workbook.worksheets.getItem("Assumptions");
  const lineItemName = getInputValue("soft-debt-name");
  const regExSoftDebt = new RegExp(modifyStr(lineItemName), "i");
  let newRange: Excel.Range | undefined;

  for (let i = 0; i < obj.rngSoftDebt.values.length; i++) {
    if (regExSoftDebt.test(obj.rngSoftDebt.values[i][0])) {
      newRange = sheetAssumptions.getRangeByIndexes(
        obj.rngSoftDebt.rowIndex + i + 1,
        obj.rngSoftDebt.columnIndex,
        1,
        1
      );
      break;
    }
  }

  if (!newRange) {
    console.error(`No matching soft debt line item found for: ${lineItemName}`);
    return;
  }

  newRange.load("values, address, formulas");
  await ctx.sync();

  const valSurplusShortfall = obj.surplusShortfall.values[0][0];
  let valSoftDebt = newRange.values[0][0];

  if (valSurplusShortfall !== 0) {
    valSoftDebt = valSoftDebt + valSurplusShortfall * -1;
    newRange.formulas = [[`=round(${valSoftDebt},0)`]];
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function maxFedCredits(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const applicableFrc = obj.applicableFrc.values[0][0];

  if (obj.eligibleBasis.values[0][0] > 2800000 / 0.09) {
    if (obj.ddaQctStatus.values[0][0] === "Y") {
      obj.volReductionFed.formulas = [[`=(${obj.eligibleBasis.address}-(2800000/1.3/${applicableFrc}/0.09))`]];
    } else {
      obj.volReductionFed.formulas = [[`=(${obj.eligibleBasis.address}-(2800000/1.0/${applicableFrc}/0.09))`]];
    }
  } else {
    console.log("Not enough eligible basis for max federal credits");
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function fedCreditAdjstr(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const stateCreditProceeds =
    obj.currStateCreditAmt.values[0][0] * obj.statePricing.values[0][0] * obj.lpProfitLoss.values[0][0];
  const lpEquityNeeded = obj.lpEquity.values[0][0] - stateCreditProceeds + -1 * obj.surplusShortfall.values[0][0];
  const applicableFrc = obj.applicableFrc.values[0][0];

  if (obj.surplusShortfall.values[0][0] < -0.5 || obj.surplusShortfall.values[0][0] > 0.5) {
    if (obj.ddaQctStatus.values[0][0] === "Y") {
      obj.volReductionFed.formulas = [[
        `=(${obj.eligibleBasis.address}-min((2800000/1.3/0.09),max((${lpEquityNeeded}/${obj.lpProfitLoss.values[0][0]}/${obj.fedPricing.values[0][0]}/10/0.09/${applicableFrc}/1.3),0)))`,
      ]];
    } else {
      obj.volReductionFed.formulas = [[
        `=(${obj.eligibleBasis.address}-min((2800000/1.0/0.09),max((${lpEquityNeeded}/${obj.lpProfitLoss.values[0][0]}/${obj.fedPricing.values[0][0]}/10/0.09/${applicableFrc}/1.0),0)))`,
      ]];
    }
  } else {
    console.log("No surplus or shortfall to bridge");
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function stateCredAdjstr(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const currStateEquity =
    obj.currStateCreditAmt.values[0][0] * obj.lpProfitLoss.values[0][0] * obj.statePricing.values[0][0];
  const stateEquityNeeded = currStateEquity + -1 * obj.surplusShortfall.values[0][0];
  const stateCreditAmtNeed =
    stateEquityNeeded /
    obj.lpProfitLoss.values[0][0] /
    obj.statePricing.values[0][0] /
    obj.stateCreditRate.values[0][0];

  if (obj.surplusShortfall.values[0][0] < -0.5 || obj.surplusShortfall.values[0][0] > 0.5) {
    obj.volReductionState.formulas = [[
      `=Min((${obj.eligibleBasis.address}-(${stateCreditAmtNeed}/1.0))*-1,${obj.volReductionFed.address})`,
    ]];
  } else {
    console.log("No surplus or shortfall to bridge");
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function deferFee(obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) {
  const existingFeeFormula = obj.cashFeeLimit.formulas[0][0] as string;

  if (obj.surplusShortfall.values[0][0] < 0) {
    const shortfall = obj.surplusShortfall.values[0][0];
    if (obj.is4Pct.values[0][0] === "Y") {
      // Strip leading = before wrapping in a new formula
      const innerFormula = existingFeeFormula.startsWith("=") ? existingFeeFormula.slice(1) : existingFeeFormula;
      obj.cashFeeLimit.formulas = [[`=${innerFormula}+${shortfall}`]];
    } else {
      const innerFormula = existingFeeFormula.startsWith("=") ? existingFeeFormula.slice(1) : existingFeeFormula;
      obj.cashFeeLimit.formulas = [[`=MAX(${innerFormula}+${shortfall}, ${obj.maxDevFee.values[0][0]}/2)`]];
    }
  } else {
    console.log("No deficit to offset");
  }

  ctx.workbook.application.iterativeCalculation.maxIteration = 30;
  ctx.workbook.application.calculate("Full");
  await ctx.sync();
}

async function tryCatch(
  callback: (fnx: (obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) => Promise<void>) => Promise<void>,
  callbackParameter: (obj: Record<string, Excel.Range>, ctx: Excel.RequestContext) => Promise<void>
) {
  try {
    await callback(callbackParameter);
  } catch (error) {
    console.error(error);
  }
}
