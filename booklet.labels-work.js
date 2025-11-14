// *** MERGED booklet.js (original + scaling/labels/modal injected) ***
(function (w) {
  'use strict';

  const CM_TO_PX = 37.7952755906;
  const MARKER_TICK_LENGTH_PX = Math.round(1 * CM_TO_PX);
  const LABEL_CLEARANCE_PX = Math.round(0.3 * CM_TO_PX);
  const DROPLINE_COLOR = 'rgba(15,23,42,0.3)';
  const DROPLINE_TICK_COLOR = 'rgba(15,23,42,0.45)';

  const isNum = (v) => typeof v === 'number' && isFinite(v);

  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

  const decs = (value) => {
    if (!isNum(value)) return 0;
    const str = String(value);
    if (!str.includes('.')) return 0;
    return str.split('.')[1].length;
  };

  const detectThousandths = (values) => values.some((v) => {
    if (!isNum(v)) return false;
    return Math.abs(v) < 1 && countDecimals(v) >= 3;
  });

  const countDecimals = (value) => {
    if (!Number.isFinite(value)) return 0;
    if (Math.abs(value) < 1e-12) return 0;
    const normalized = Number(value.toPrecision(12));
    if (!Number.isFinite(normalized)) return 0;
    const str = String(normalized).toLowerCase();
    if (str.includes('e')) {
      const [base, exponent] = str.split('e');
      const decimalsPart = (base.split('.')[1] || '').replace(/0+$/, '');
      const exp = Number(exponent);
      return Math.max(0, decimalsPart.length - exp);
    }

    if (!str.includes('.')) return 0;
    return str.split('.')[1].replace(/0+$/, '').length;
  };

  const roundToDecimals = (value, decimals) => {
    const safeDecimals = Math.max(0, Math.min(decimals, 8));
    const factor = Math.pow(10, safeDecimals + 2);
    return Math.round(value * factor) / factor;
  };

  const niceStep = (rawStep, constraints = {}) => {
    const allowDecimalTicks = !!constraints.allowDecimalTicks;
    const minimum = allowDecimalTicks
      ? Math.max(Number.EPSILON, constraints.minimum || Number.EPSILON)
      : Math.max(1, constraints.minimum || 1);
    if (!Number.isFinite(rawStep) || rawStep === 0) return minimum;
    const absStep = Math.abs(rawStep);
    const exponent = Math.floor(Math.log10(absStep));
    const magnitude = Math.pow(10, exponent);
    const fraction = absStep / magnitude;
    let niceFraction;
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2.5) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
    let step = niceFraction * magnitude;
    if (constraints.isPercent) {
      const percentChoices = allowDecimalTicks
        ? [0.001, 0.002, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.25, 0.5, 1, 2, 5, 10, 20, 25, 50]
        : [1, 2, 5, 10, 20, 25, 50];
      const adjusted = percentChoices.find((choice) => choice >= step) || percentChoices[percentChoices.length - 1];
      step = Math.max(adjusted, minimum);
    } else {
      step = Math.max(step, minimum);
    }
    if (!allowDecimalTicks && !Number.isInteger(step)) {
      step = Math.ceil(step);
    }
    return rawStep < 0 ? -step : step;
  };

  function computeAxisBounds(stat, options = {}) {
  const debugLog = Array.isArray(options?.__debugCollector) ? options.__debugCollector : null;
  const seriesList = Array.isArray(stat?.series) ? stat.series : [];
    const desiredTickCountInput = Number(options.desiredTickCount);
    const desiredTickCount = Number.isFinite(desiredTickCountInput)
      ? Math.min(Math.max(desiredTickCountInput, 4), 12)
      : 6;
    const targetDivisions = Math.max(4, Math.min(12, desiredTickCount - 1));
    const decimalCapOption = Number.isFinite(options.decimalCap)
      ? Math.max(0, Math.min(options.decimalCap, 6))
      : null;
    const padFraction = Number.isFinite(options.padFraction)
      ? Math.max(0, Math.min(options.padFraction, 0.4))
      : (stat.isPercent ? 0.03 : 0.08);
    const requestedMinInterval = Number.isFinite(options.minimumTickInterval)
      ? Math.max(Math.abs(options.minimumTickInterval), Number.EPSILON)
      : null;

    const overrideMin = Number.isFinite(options.min) ? Number(options.min) : null;
    const overrideMax = Number.isFinite(options.max) ? Number(options.max) : null;
    const hasOverrideMin = overrideMin != null;
    const hasOverrideMax = overrideMax != null;
    const manualSegments = hasOverrideMin && hasOverrideMax && overrideMax > overrideMin
      ? Math.min(Math.max(Number(options.manualDivisions) || (desiredTickCount - 1), 4), 16)
      : null;

    const rawExplicitMin = Number.isFinite(stat?.min) ? Number(stat.min) : null;
    const rawExplicitMax = Number.isFinite(stat?.max) ? Number(stat.max) : null;
    const metaExplicitMin = seriesList.reduce((acc, series) => {
      const candidate = Number.isFinite(series?.meta?.min) ? Number(series.meta.min) : null;
      if (candidate == null) return acc;
      return acc == null ? candidate : Math.min(acc, candidate);
    }, null);
    const metaExplicitMax = seriesList.reduce((acc, series) => {
      const candidate = Number.isFinite(series?.meta?.max) ? Number(series.meta.max) : null;
      if (candidate == null) return acc;
      return acc == null ? candidate : Math.max(acc, candidate);
    }, null);
    const explicitMinSource = rawExplicitMin != null ? rawExplicitMin : metaExplicitMin;
    const explicitMaxSource = rawExplicitMax != null ? rawExplicitMax : metaExplicitMax;

    const values = [];
    let maxDecimalPrecision = 0;
    seriesList.forEach((series) => {
      const data = Array.isArray(series?.data) ? series.data : [];
      data.forEach((value) => {
        if (Number.isFinite(value)) {
          values.push(value);
          maxDecimalPrecision = Math.max(maxDecimalPrecision, countDecimals(value));
        }
      });
    });

  const ZERO_EPSILON = 1e-9;
  const hasValues = values.length > 0;
  const hasThousandths = hasValues && detectThousandths(values);
  const percentWithinBounds = !!stat?.isPercent && (!hasValues || values.every((value) => value >= -ZERO_EPSILON && value <= 100 + ZERO_EPSILON));
  const explicitRange = explicitMinSource != null && explicitMaxSource != null && explicitMaxSource > explicitMinSource
    ? explicitMaxSource - explicitMinSource
    : null;
  const hasNegatives = hasValues && values.some((value) => value < -ZERO_EPSILON);
    const hasPositives = hasValues && values.some((value) => value > ZERO_EPSILON);
    const allZero = hasValues && !hasNegatives && !hasPositives;

    let dataMin = hasValues ? Math.min(...values) : (stat.isPercent ? 0 : 0);
    let dataMax = hasValues ? Math.max(...values) : (stat.isPercent ? 100 : 1);

  const explicitMin = explicitMinSource != null && hasNegatives ? explicitMinSource : (explicitMinSource != null ? Math.max(0, explicitMinSource) : null);
  const explicitMax = explicitMaxSource;
  const boundMin = hasOverrideMin ? overrideMin : null;
  const boundMax = hasOverrideMax ? overrideMax : null;
  if (explicitMin != null) dataMin = Math.min(dataMin, explicitMin);
  if (explicitMax != null) dataMax = Math.max(dataMax, explicitMax);
    if (hasOverrideMin) dataMin = Math.min(dataMin, overrideMin);
    if (hasOverrideMax) dataMax = Math.max(dataMax, overrideMax);

  const percentHighCluster = percentWithinBounds && !hasNegatives && dataMin >= 90 && dataMax <= 100;
    const percentHighClusterIntegers = percentHighCluster && maxDecimalPrecision === 0 && !hasThousandths;
    if (percentHighClusterIntegers) {
      dataMin = Math.min(dataMin, 70);
      dataMax = Math.max(dataMax, 100);
    }

    if (allZero && !stat.isPercent) {
      if (boundMin != null) {
        dataMin = boundMin;
      } else {
        dataMin = 0;
      }
      if (boundMax != null) {
        dataMax = boundMax;
      } else if (!hasOverrideMin && !hasOverrideMax && explicitMax == null) {
        dataMax = Math.max(dataMax, 15);
      } else {
        dataMax = Math.max(dataMax, dataMin === 0 ? 1 : dataMin + 1);
      }
    } else if (dataMin === dataMax) {
      const lift = Math.max(Math.abs(dataMax || 1) * 0.25, stat.isPercent ? 1 : 0.5);
      dataMin -= lift;
      dataMax += lift;
    }

  let paddedMin = hasOverrideMin ? overrideMin : dataMin;
  let paddedMax = hasOverrideMax ? overrideMax : dataMax;
    const span = Math.max(paddedMax - paddedMin, Number.EPSILON);
    let paddingValue = span * padFraction;
    if (!Number.isFinite(paddingValue) || paddingValue <= ZERO_EPSILON) {
      paddingValue = (stat.isPercent ? 1 : Math.max(Math.abs(paddedMax) || 1, 1)) * padFraction;
    }

  if ((!hasOverrideMin || !hasOverrideMax) && hasValues) {
    const rawSpan = Math.max(Math.abs(dataMax - dataMin), Number.EPSILON);
    if (Number.isFinite(rawSpan) && rawSpan > ZERO_EPSILON) {
      const stepSeed = rawSpan / Math.max(targetDivisions, 1);
      if (Number.isFinite(stepSeed) && stepSeed > ZERO_EPSILON) {
        const padMultiplier = stat.isPercent ? 0.35 : 0.5;
        const minimumPad = Math.max(stepSeed * padMultiplier, Number.EPSILON);
        const cappedPad = (stat.isPercent && percentWithinBounds)
          ? Math.min(minimumPad, 10)
          : minimumPad;
        paddingValue = Math.max(paddingValue, cappedPad);
      }
    }
  }

  const padBoost = span * (stat.isPercent ? 0.005 : 0.01);

  if (!hasOverrideMin) paddedMin -= paddingValue + padBoost;
  if (!hasOverrideMax) paddedMax += paddingValue + padBoost;


    if (stat.isPercent && percentWithinBounds) {
      paddedMin = Math.max(0, paddedMin);
      paddedMax = Math.min(100, paddedMax);
    } else if (!hasNegatives && !hasOverrideMin) {
      paddedMin = Math.max(0, paddedMin);
      if (!stat.isPercent && !hasOverrideMax) {
        paddedMax = Math.max(paddedMax, paddedMin === 0 && allZero ? 15 : paddedMax);
      }
    }

    if (allZero && !stat.isPercent && !hasOverrideMax && explicitMax == null) {
      paddedMax = Math.max(paddedMax, 15);
    }

    if (paddedMax <= paddedMin) {
      const adjustment = Math.max(Math.abs(paddedMin) || 1, 1);
      paddedMax = paddedMin + adjustment;
    }

  const explicitSegment = explicitRange != null ? Math.abs(explicitRange) / Math.max(targetDivisions, 1) : null;
  const overrideRange = hasOverrideMin && hasOverrideMax ? Math.abs(overrideMax - overrideMin) : null;
      const explicitSegmentPrecision = Number.isFinite(explicitSegment) ? countDecimals(explicitSegment) : 0;
      const explicitMinHasFraction = explicitMin != null && Math.abs(explicitMin - Math.round(explicitMin)) > ZERO_EPSILON;
      const explicitMaxHasFraction = explicitMax != null && Math.abs(explicitMax - Math.round(explicitMax)) > ZERO_EPSILON;
      const overrideMinHasFraction = hasOverrideMin && Math.abs(overrideMin - Math.round(overrideMin)) > ZERO_EPSILON;
      const overrideMaxHasFraction = hasOverrideMax && Math.abs(overrideMax - Math.round(overrideMax)) > ZERO_EPSILON;
      const overridePrecision = Math.max(
        overrideMinHasFraction ? countDecimals(overrideMin) : 0,
        overrideMaxHasFraction ? countDecimals(overrideMax) : 0
      );
      const fractionalEvidence = maxDecimalPrecision > 0
          || explicitMinHasFraction
          || explicitMaxHasFraction
          || overrideMinHasFraction
          || overrideMaxHasFraction
          || hasThousandths;
        const baseDecimalCap = stat.isPercent ? 2 : 4;
        const autoDecimalCap = Math.min(6, Math.max(baseDecimalCap, maxDecimalPrecision, explicitSegmentPrecision, hasThousandths ? 3 : 0, overridePrecision));
    let decimalCap = decimalCapOption != null ? decimalCapOption : autoDecimalCap;
        const spanMagnitude = Math.abs(paddedMax - paddedMin);
        const dominantMagnitude = Math.max(
          Math.abs(paddedMin),
          Math.abs(paddedMax),
          Math.abs(dataMin),
          Math.abs(dataMax)
        );
      const manualStepEstimate = manualSegments != null && Number.isFinite(overrideRange)
        ? Math.abs(overrideRange) / Math.max(manualSegments, 1)
        : null;
      const estimatedAutoStep = Number.isFinite(spanMagnitude)
        ? spanMagnitude / Math.max(targetDivisions, 1)
        : null;
      const allowFineDecimals = hasThousandths
        || dominantMagnitude < 1 + ZERO_EPSILON
        || spanMagnitude < 1 - ZERO_EPSILON
        || (manualStepEstimate != null && manualStepEstimate < 1 - ZERO_EPSILON);
        if (stat.isPercent) {
          const percentMinDecimals = allowFineDecimals ? 2 : 0;
          const percentMaxDecimals = allowFineDecimals ? 4 : 1;
          if (manualStepEstimate != null && manualStepEstimate > ZERO_EPSILON && manualStepEstimate < 1 - ZERO_EPSILON) {
            const manualStepDecimals = Math.max(0, countDecimals(manualStepEstimate));
            const targetDecimals = Math.min(percentMaxDecimals, Math.max(percentMinDecimals, manualStepDecimals));
            decimalCap = Math.max(decimalCap, targetDecimals);
          }
          decimalCap = clamp(decimalCap, percentMinDecimals, percentMaxDecimals);
        } else if (!allowFineDecimals) {
          decimalCap = Math.min(decimalCap, dominantMagnitude < 10 ? 2 : 1);
        }
      let allowDecimalTicks = fractionalEvidence
        || (requestedMinInterval && requestedMinInterval < 1)
        || (Number.isFinite(estimatedAutoStep) && estimatedAutoStep < 1 - ZERO_EPSILON)
        || (Number.isFinite(manualStepEstimate) && manualStepEstimate < 1 - ZERO_EPSILON);
      const spanToMagnitudeRatio = dominantMagnitude > ZERO_EPSILON
        ? spanMagnitude / dominantMagnitude
        : Number.POSITIVE_INFINITY;
      if (!manualSegments
        && !stat.isPercent
        && !hasThousandths
        && !fractionalEvidence
        && dominantMagnitude >= 100 - ZERO_EPSILON
        && spanToMagnitudeRatio < 0.05 - ZERO_EPSILON) {
        allowDecimalTicks = false;
      }
    const baseMinimumTickInterval = allowDecimalTicks
      ? Math.pow(10, -decimalCap)
      : 1;
    let minimumTickInterval = baseMinimumTickInterval;
    if (requestedMinInterval != null) {
      if (allowDecimalTicks) {
        minimumTickInterval = Math.max(Number.EPSILON, Math.min(requestedMinInterval, baseMinimumTickInterval));
      } else {
        minimumTickInterval = Math.max(requestedMinInterval, baseMinimumTickInterval);
      }
    }

    if (manualStepEstimate != null && manualStepEstimate > ZERO_EPSILON) {
      minimumTickInterval = Math.min(minimumTickInterval, Math.max(manualStepEstimate, Number.EPSILON));
    }

    if (allowDecimalTicks) {
      if (stat.isPercent) {
        if (manualStepEstimate != null && manualStepEstimate > ZERO_EPSILON && manualStepEstimate < 0.1) {
          minimumTickInterval = Math.max(Number.EPSILON, Math.min(minimumTickInterval, manualStepEstimate));
        } else if (allowFineDecimals) {
          minimumTickInterval = Math.max(minimumTickInterval, 0.1);
        } else {
          minimumTickInterval = Math.max(minimumTickInterval, 1);
        }
      } else if (!allowFineDecimals) {
        minimumTickInterval = Math.max(minimumTickInterval, 0.1);
      }
    }

    if (allowDecimalTicks && !stat.isPercent) {
      minimumTickInterval = Math.max(minimumTickInterval, 0.01);
    }

    const computeNiceStep = (range, divisions) => {
      const denominator = Math.max(divisions, 1);
      const raw = range / denominator;
      if (!Number.isFinite(raw) || raw <= 0) return minimumTickInterval;
      return Math.max(raw, minimumTickInterval);
    };

    const snapStep = (value) => {
      if (!Number.isFinite(value) || value <= 0) return minimumTickInterval;
      const absValue = Math.max(Math.abs(value), minimumTickInterval);
      const exponent = Math.floor(Math.log10(absValue));
      const magnitude = Math.pow(10, exponent);
      const baseDigits = allowDecimalTicks
        ? [1, 1.25, 1.5, 2, 2.5, 3, 4, 5]
        : [1, 2, 3, 4, 5];
      const candidates = new Set();
      baseDigits.forEach((digit) => {
        candidates.add(digit * magnitude);
        candidates.add(digit * magnitude * 10);
        candidates.add(digit * magnitude / 10);
      });
      let snapped = null;
      let bestDiff = Number.POSITIVE_INFINITY;
      candidates.forEach((candidate) => {
        if (!Number.isFinite(candidate)) return;
        const adjusted = Math.max(candidate, minimumTickInterval);
        const diff = Math.abs(adjusted - absValue);
        if (diff < bestDiff - 1e-12 || (Math.abs(diff - bestDiff) <= 1e-12 && (snapped == null || adjusted < snapped))) {
          bestDiff = diff;
          snapped = adjusted;
        }
      });
      if (snapped == null) snapped = Math.max(absValue, minimumTickInterval);
      if (!allowDecimalTicks) {
        snapped = Math.max(1, Math.round(snapped));
      }
      return Math.max(snapped, minimumTickInterval);
    };

    const MIN_AUTO_SEGMENTS = 4;
    const MAX_AUTO_SEGMENTS = 15;
    const preferredSegments = manualSegments
      ? manualSegments
      : Math.max(MIN_AUTO_SEGMENTS, Math.min(targetDivisions, MAX_AUTO_SEGMENTS));
    const rangeForSteps = Math.max(paddedMax - paddedMin, minimumTickInterval);
    const baseStepRaw = manualSegments
      ? Math.max(Math.abs(overrideMax - overrideMin) / manualSegments, minimumTickInterval)
      : rangeForSteps / Math.max(preferredSegments, 1);

    const minimumSegments = manualSegments || MIN_AUTO_SEGMENTS;
    const maximumSegments = manualSegments || MAX_AUTO_SEGMENTS;

  const candidateSteps = new Set();
    const addCandidateStep = (value, { snap = true } = {}) => {
      if (!Number.isFinite(value)) return;
      const baseValue = snap ? snapStep(value) : Math.max(value, minimumTickInterval);
      const stepValue = allowDecimalTicks ? Math.max(baseValue, minimumTickInterval) : Math.max(1, Math.round(baseValue));
      if (!Number.isFinite(stepValue) || stepValue <= 0) return;
      const normalized = Number(stepValue.toPrecision(12));
      if (normalized > 0) candidateSteps.add(normalized);
      if (debugLog) debugLog.push({ phase: 'candidate', raw: value, snapped: normalized });
    };

    if (manualSegments) {
      addCandidateStep(baseStepRaw, { snap: false });
    } else {
      const niceBase = snapStep(baseStepRaw);
      addCandidateStep(niceBase);
      addCandidateStep(baseStepRaw);

      const clampSegments = (value) => Math.max(minimumSegments, Math.min(value, maximumSegments));
      const segmentSeeds = new Set([
        clampSegments(preferredSegments),
        clampSegments(preferredSegments + 1),
        clampSegments(preferredSegments - 1),
        minimumSegments,
        clampSegments(minimumSegments + 1),
        clampSegments(minimumSegments + 2),
        clampSegments(Math.round(preferredSegments * 1.5)),
        clampSegments(Math.round(preferredSegments * 2))
      ]);

      if (explicitRange != null && explicitRange > 0) {
        const explicitSegments = Math.max(1, Math.round(explicitRange / Math.max(minimumTickInterval, niceBase)));
        segmentSeeds.add(clampSegments(explicitSegments));
      }

      segmentSeeds.forEach((segmentsCandidate) => {
        if (!Number.isFinite(segmentsCandidate) || segmentsCandidate <= 0) return;
        const raw = rangeForSteps / segmentsCandidate;
        addCandidateStep(raw);
      });

      if (allowDecimalTicks) {
        addCandidateStep(niceBase / 2);
        addCandidateStep(niceBase / 5);
        addCandidateStep(niceBase * 2);
      } else {
        addCandidateStep(niceBase * 2);
        addCandidateStep(niceBase / 2);
      }

      addCandidateStep(minimumTickInterval);
    }

    if (boundMin != null && boundMax != null && boundMax > boundMin) {
      for (let segmentsCandidate = minimumSegments; segmentsCandidate <= maximumSegments; segmentsCandidate += 1) {
        const boundedStep = (boundMax - boundMin) / segmentsCandidate;
        addCandidateStep(boundedStep);
      }
    }

    const evaluateStep = (stepValue) => {
      if (!Number.isFinite(stepValue) || stepValue <= 0) return null;
      const step = allowDecimalTicks ? stepValue : Math.max(1, Math.round(stepValue));
      const stepDecimals = allowDecimalTicks
        ? Math.min(decimalCap, Math.max(countDecimals(step), maxDecimalPrecision, explicitSegmentPrecision, overridePrecision))
        : 0;
      const roundingFactor = Math.pow(10, stepDecimals);
      const roundValue = (value) => Math.round(value * roundingFactor) / roundingFactor;
      const floorToStep = (value) => roundValue(Math.floor((value + ZERO_EPSILON) / step) * step);
      const ceilToStep = (value) => roundValue(Math.ceil((value - ZERO_EPSILON) / step) * step);

      let axisMin = hasOverrideMin ? roundValue(overrideMin) : floorToStep(paddedMin);
      let axisMax = hasOverrideMax ? roundValue(overrideMax) : ceilToStep(paddedMax);

      if (manualSegments) {
        axisMin = roundValue(overrideMin);
        axisMax = roundValue(overrideMax);
      }

      if (stat.isPercent) {
        axisMin = Math.max(0, axisMin);
        if (percentWithinBounds) {
          axisMax = Math.min(100, axisMax);
        }
      } else {
        if (!hasNegatives && !hasOverrideMin) {
          axisMin = Math.max(0, axisMin);
        }
        if (!hasNegatives && !hasOverrideMax) {
          axisMax = Math.max(axisMax, 0);
        }
      }

      if (!hasOverrideMin) axisMin = floorToStep(axisMin);
      if (!hasOverrideMax) {
        axisMax = ceilToStep(axisMax);
        if (boundMax != null && dataMax <= boundMax + ZERO_EPSILON) {
          axisMax = Math.min(axisMax, roundValue(boundMax));
        }
        if (stat.isPercent && percentWithinBounds && axisMax > 100 + 1e-9) {
          let adjustedMax = axisMax;
          while (adjustedMax > 100 + 1e-9 && adjustedMax - step >= paddedMax - 1e-9) {
            adjustedMax = roundValue(adjustedMax - step);
          }
          if (adjustedMax > 100 + 1e-9) {
            return null;
          }
          axisMax = Math.min(adjustedMax, 100);
        }
      }

      if (boundMin != null && dataMin >= boundMin - ZERO_EPSILON) {
        axisMin = Math.max(axisMin, roundValue(boundMin));
      }
      if (boundMax != null && dataMax <= boundMax + ZERO_EPSILON) {
        axisMax = Math.min(axisMax, roundValue(boundMax));
      }

      if (axisMax <= axisMin) {
        axisMax = roundValue(axisMin + step);
      }

      const spanStepCount = (axisMax - axisMin) / step;
      if (!manualSegments && (spanStepCount <= 0 || Math.abs(spanStepCount - Math.round(spanStepCount)) > 1e-6)) {
        return null;
      }

      if (!manualSegments) {
        const desiredSpan = step * minimumSegments;
        const canExtendBelowZero = hasNegatives || (stat.isPercent && !percentWithinBounds);
        const lowerBound = hasOverrideMin
          ? axisMin
          : (boundMin != null ? roundValue(boundMin) : (canExtendBelowZero ? -Number.POSITIVE_INFINITY : 0));
        const upperBound = hasOverrideMax
          ? axisMax
          : (boundMax != null ? roundValue(boundMax) : (stat.isPercent && percentWithinBounds ? 100 : Number.POSITIVE_INFINITY));
        let guardExpand = 0;
        while ((axisMax - axisMin) < (desiredSpan - 1e-9) && guardExpand < 128) {
          let grew = false;
          if (!hasOverrideMin && axisMin - step >= lowerBound - 1e-9) {
            axisMin = roundValue(axisMin - step);
            grew = true;
          }
          if ((axisMax - axisMin) >= (desiredSpan - 1e-9)) break;
          if (!hasOverrideMax && axisMax + step <= upperBound + 1e-9) {
            axisMax = roundValue(axisMax + step);
            grew = true;
          }
          if (!grew) break;
          guardExpand += 1;
        }
      }

      const ticks = [];
      let guard = 0;
      let current = axisMin;
      while (current <= axisMax + step * 0.5 && guard < 256) {
        ticks.push(roundValue(current));
        current = roundValue(current + step);
        guard += 1;
      }

      if (!ticks.includes(axisMax)) ticks.push(roundValue(axisMax));
      if (!ticks.includes(axisMin)) ticks.unshift(roundValue(axisMin));

      let segments = ticks.length - 1;
      if (manualSegments && segments > manualSegments) {
        ticks.length = manualSegments + 1;
        segments = manualSegments;
      }

      if (segments < minimumSegments || segments > maximumSegments) {
        if (debugLog) debugLog.push({ phase: 'reject', reason: 'segment-count', step, segments, minimumSegments, maximumSegments });
        return null;
      }

      const ensureAligned = (value) => {
        if (!Number.isFinite(value)) return;
        const normalized = roundValue(value);
        const offset = Math.abs((normalized - ticks[0]) / step - Math.round((normalized - ticks[0]) / step));
        if (offset > 1e-6) return;
        if (ticks.includes(normalized)) return;

        if (normalized > ticks[ticks.length - 1]) {
          let cursor = roundValue(ticks[ticks.length - 1] + step);
          let guardFill = 0;
          while (cursor < normalized - step * 0.5 && guardFill < 128) {
            ticks.push(cursor);
            cursor = roundValue(cursor + step);
            guardFill += 1;
          }
          ticks.push(normalized);
          return;
        }

        if (normalized < ticks[0]) {
          let cursor = roundValue(ticks[0] - step);
          let guardFill = 0;
          while (cursor > normalized + step * 0.5 && guardFill < 128) {
            ticks.push(cursor);
            cursor = roundValue(cursor - step);
            guardFill += 1;
          }
          ticks.push(normalized);
          return;
        }

        ticks.push(normalized);
      };

      if (ticks[0] <= 0 && ticks[ticks.length - 1] >= 0) ensureAligned(0);
      if (stat.isPercent && percentWithinBounds) ensureAligned(100);
      ensureAligned(explicitMin);
      ensureAligned(explicitMax);

      ticks.sort((a, b) => a - b);
      const uniqueTicks = Array.from(new Set(ticks)).sort((a, b) => a - b);
      const finalSegments = uniqueTicks.length - 1;

      if (finalSegments < minimumSegments || finalSegments > maximumSegments) {
        return null;
      }

      const finalMin = uniqueTicks[0];
      const finalMax = uniqueTicks[uniqueTicks.length - 1];
    const coverageLower = Math.max(0, paddedMin - finalMin);
    const coverageUpper = Math.max(0, finalMax - paddedMax);
    const softCoverageTarget = Math.max(paddingValue, minimumTickInterval || 0);
    const coverageMissLower = !hasOverrideMin ? Math.max(0, softCoverageTarget - coverageLower) : 0;
    const coverageMissUpper = !hasOverrideMax ? Math.max(0, softCoverageTarget - coverageUpper) : 0;
  const segmentDiff = Math.abs(finalSegments - preferredSegments);
    const stepDrift = Math.abs(step - baseStepRaw) / Math.max(step, baseStepRaw, 1);
      const tickPrecision = uniqueTicks.reduce((acc, value) => Math.max(acc, countDecimals(value)), 0);
      const stepPrecision = countDecimals(step);
      const minIntervalPrecision = requestedMinInterval ? countDecimals(requestedMinInterval) : 0;
      const precisionPool = [tickPrecision, stepPrecision, maxDecimalPrecision, overridePrecision, minIntervalPrecision];
      if (Number.isFinite(explicitSegment) && Math.abs(explicitSegment - step) <= Math.max(Math.abs(step), Math.abs(explicitSegment)) * 1e-6) {
        precisionPool.push(explicitSegmentPrecision);
      }
      const rawDisplayDecimals = allowDecimalTicks
        ? Math.min(decimalCap, Math.max(...precisionPool))
        : 0;
      let displayDecimals = rawDisplayDecimals;
      if (allowDecimalTicks) {
        const axisStepLimit = step < 0.1 - ZERO_EPSILON ? 2 : 1;
        displayDecimals = Math.min(displayDecimals, axisStepLimit);
      }

      const evenDivision = manualSegments
        ? Math.abs(spanStepCount - manualSegments) < 1e-6
        : Math.abs(spanStepCount - Math.round(spanStepCount)) < 1e-6;
      const coveragePenalty = (coverageMissLower + coverageMissUpper) * 1.5e6;
      const score = (segmentDiff * 2e6)
        + ((coverageLower + coverageUpper) * 4e5)
        + coveragePenalty
        + (displayDecimals * 5e2)
        + (stepDrift * 5e3)
        + (evenDivision ? 0 : 1e7);

      if (debugLog) {
        debugLog.push({
          phase: 'decimals',
          step,
          displayDecimals,
          tickPrecision,
          decimalCap,
          maxDecimalPrecision,
          explicitSegmentPrecision,
          overridePrecision
        });
      }

      const outcome = {
        min: finalMin,
        max: finalMax,
        ticks: uniqueTicks,
        decimals: rawDisplayDecimals,
        displayDecimals,
        allowDecimals: allowDecimalTicks,
        score
      };

      if (debugLog) debugLog.push({ phase: 'accept', step, segments: finalSegments, min: finalMin, max: finalMax, score, coverageUpper, coverageLower });
      return outcome;
    };

    let bestCandidate = null;
    candidateSteps.forEach((stepValue) => {
      const evaluated = evaluateStep(stepValue);
      if (!evaluated) return;
      if (!bestCandidate || evaluated.score < bestCandidate.score) {
        bestCandidate = evaluated;
      }
    });

    if (!bestCandidate) {
      bestCandidate = evaluateStep(minimumTickInterval);
    }

    if (!bestCandidate) {
      if (manualSegments && hasOverrideMin && hasOverrideMax) {
        const manualSpan = Math.max(Math.abs(overrideMax - overrideMin), minimumTickInterval);
        const manualStepRaw = manualSpan / Math.max(manualSegments, 1);
        const manualStep = allowDecimalTicks
          ? Math.max(manualStepRaw, minimumTickInterval)
          : Math.max(1, Math.round(manualStepRaw));
        let manualDecimals = allowDecimalTicks
          ? Math.min(decimalCap, Math.max(countDecimals(manualStep), maxDecimalPrecision, explicitSegmentPrecision, overridePrecision))
          : 0;
        if (allowDecimalTicks) {
          if (!allowFineDecimals) {
            manualDecimals = Math.min(manualDecimals, stat.isPercent ? 1 : (dominantMagnitude < 10 ? 2 : 1));
          } else if (stat.isPercent) {
            manualDecimals = Math.min(manualDecimals, 1);
          }
        }
        const roundingFactor = Math.pow(10, manualDecimals);
        const roundValue = (value) => Math.round(value * roundingFactor) / roundingFactor;
        const ticks = [];
        for (let i = 0; i <= manualSegments; i += 1) {
          const value = overrideMin + manualStep * i;
          ticks.push(roundValue(value));
        }
        const uniqueTicks = Array.from(new Set(ticks)).sort((a, b) => a - b);
        const finalMin = roundValue(Math.min(overrideMin, overrideMax));
        const finalMax = roundValue(Math.max(overrideMin, overrideMax));

        return {
          min: uniqueTicks.length ? uniqueTicks[0] : finalMin,
          max: uniqueTicks.length ? uniqueTicks[uniqueTicks.length - 1] : finalMax,
          ticks: uniqueTicks.length ? uniqueTicks : [finalMin, finalMax],
          decimals: manualDecimals,
          displayDecimals: Math.min(manualDecimals, stat && stat.isPercent ? 1 : (dominantMagnitude < 10 ? 2 : manualDecimals)),
          allowDecimals: allowDecimalTicks,
          reversed: !!stat?.upsideDown
        };
      }

      const fallbackMinSource = hasOverrideMin ? overrideMin : paddedMin;
      const fallbackMaxSource = hasOverrideMax ? overrideMax : paddedMax;
      const fallbackRange = Math.max(fallbackMaxSource - fallbackMinSource, minimumTickInterval);
      const fallbackSegments = manualSegments || Math.max(MIN_AUTO_SEGMENTS, Math.min(preferredSegments, MAX_AUTO_SEGMENTS));
      const fallbackStepRaw = fallbackRange / Math.max(fallbackSegments, 1);
      const fallbackStep = allowDecimalTicks ? snapStep(fallbackStepRaw) : Math.max(1, Math.round(fallbackStepRaw));
      let fallbackDecimals = allowDecimalTicks
        ? Math.min(decimalCap, Math.max(countDecimals(fallbackStep), maxDecimalPrecision, explicitSegmentPrecision, overridePrecision))
        : 0;
      if (allowDecimalTicks) {
        if (!allowFineDecimals) {
          fallbackDecimals = Math.min(fallbackDecimals, stat.isPercent ? 1 : (dominantMagnitude < 10 ? 2 : 1));
        } else if (stat.isPercent) {
          fallbackDecimals = Math.min(fallbackDecimals, 1);
        }
      }
      const roundingFactor = Math.pow(10, fallbackDecimals);
      const roundValue = (value) => Math.round(value * roundingFactor) / roundingFactor;

      let axisMin = hasOverrideMin
        ? roundValue(overrideMin)
        : roundValue(Math.floor((fallbackMinSource + ZERO_EPSILON) / fallbackStep) * fallbackStep);
      let axisMax = hasOverrideMax
        ? roundValue(overrideMax)
        : roundValue(Math.ceil((fallbackMaxSource - ZERO_EPSILON) / fallbackStep) * fallbackStep);

      if (stat.isPercent && percentWithinBounds) {
        axisMin = Math.max(0, axisMin);
        axisMax = Math.min(100, axisMax);
      } else if (!hasNegatives && !hasOverrideMin) {
        axisMin = Math.max(0, axisMin);
      }

      if (boundMin != null && dataMin >= boundMin - ZERO_EPSILON) {
        axisMin = Math.max(axisMin, roundValue(boundMin));
      }
      if (boundMax != null && dataMax <= boundMax + ZERO_EPSILON) {
        axisMax = Math.min(axisMax, roundValue(boundMax));
      }

      if (axisMax <= axisMin) {
        axisMax = roundValue(axisMin + fallbackStep);
      }

      const ticks = [];
      let guard = 0;
      let cursor = axisMin;
      while (cursor <= axisMax + fallbackStep * 0.5 && guard < 256) {
        ticks.push(roundValue(cursor));
        cursor += fallbackStep;
        guard += 1;
      }
      if (!ticks.includes(axisMax)) {
        ticks.push(axisMax);
      }

      const uniqueTicks = Array.from(new Set(ticks)).sort((a, b) => a - b);
      const targetSegments = manualSegments || MIN_AUTO_SEGMENTS;
      while (uniqueTicks.length - 1 < targetSegments) {
        const last = uniqueTicks[uniqueTicks.length - 1];
        const nextTick = roundValue(last + fallbackStep);
        if (uniqueTicks.includes(nextTick)) {
          break;
        }
        uniqueTicks.push(nextTick);
      }

      return {
        min: uniqueTicks[0],
        max: uniqueTicks[uniqueTicks.length - 1],
        ticks: uniqueTicks,
        decimals: fallbackDecimals,
        displayDecimals: Math.min(fallbackDecimals, stat && stat.isPercent ? 1 : (dominantMagnitude < 10 ? 2 : fallbackDecimals)),
        allowDecimals: allowDecimalTicks,
        reversed: !!stat?.upsideDown
      };
    }

    return {
      min: bestCandidate.min,
      max: bestCandidate.max,
      ticks: bestCandidate.ticks,
      decimals: bestCandidate.decimals,
      displayDecimals: bestCandidate.displayDecimals != null ? bestCandidate.displayDecimals : bestCandidate.decimals,
      allowDecimals: allowDecimalTicks,
      reversed: !!stat?.upsideDown
    };
  }

  function ensureHalo(chart) {
    try {
      const defs = chart.renderer.defs && chart.renderer.defs.element;
      if (defs && !defs.querySelector('#softWhiteHalo')) {
        defs.innerHTML += '<filter id="softWhiteHalo" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur in="SourceAlpha" stdDeviation="2.2" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>';
      }
    } catch (e) {
      /* noop */
    }
  }

  function enhanceLabels(chart) {
    if (!chart || !chart.series) return;
    (chart.series || []).forEach((series) => {
      if (!series || !series.points) return;
      series.points.forEach((point) => {
        const dataLabel = point && point.dataLabel;
        if (!dataLabel || typeof dataLabel.css !== 'function') return;
        dataLabel.css({ filter: 'none', textOutline: '', fontWeight: '', color: '' });
      });
    });
  }

  function installDroplines(chart) {
    if (!chart) return;
    const storeKey = '_droplineGraphics';
    const existing = chart[storeKey];
    if (Array.isArray(existing)) {
      existing.forEach((graphic) => {
        try {
          if (graphic && typeof graphic.destroy === 'function') {
            graphic.destroy();
          }
        } catch (e) {
          /* noop */
        }
      });
    }
    chart[storeKey] = [];
  }

  function onRender() {
    installDroplines(this);
    enhanceLabels(this);
  }

  if (typeof w.Highcharts !== 'undefined' && w.Highcharts.addEvent) {
    w.Highcharts.addEvent(w.Highcharts.Chart, 'render', onRender);
  }

  w.computeAxisBounds = computeAxisBounds;

})(window);


/* ===== Injected: Per-chart scale UI (modal inside container) & CSV-load spinner ===== */
(function(){
  'use strict';
  if (typeof Highcharts === 'undefined') return;

  function ensureStyles() {
    if (document.getElementById('hc-scale-ui-styles')) return;
    const css = `
      .scale-btn{position:absolute;top:8px;right:8px;width:28px;height:28px;border:none;border-radius:50%;background:rgba(0,0,0,.55);color:#fff;font-size:16px;line-height:1;display:flex;align-items:center;justify-content:center;cursor:pointer;z-index:10;opacity:0;transition:opacity .25s ease,background .2s ease}
      .scale-btn.show{opacity:1}
      .scale-btn:hover{background:rgba(0,0,0,.85)}
      .hc-loading-overlay{position:absolute;inset:0;background:rgba(0,0,0,.35);display:none;align-items:center;justify-content:center;z-index:20}
      .hc-loading-overlay.show{display:flex}
      .hc-toast-container{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:1080}
      .hc-chart-modal .modal-dialog{max-width:720px;margin:1.75rem auto}
  .hc-chart-modal .modal-content{display:flex;flex-direction:column;background:linear-gradient(160deg,rgba(11,18,32,0.95),rgba(15,23,42,0.92));color:#f8fafc;border-radius:1.1rem;border:1px solid rgba(148,163,184,0.2);box-shadow:0 32px 72px rgba(15,23,42,0.6)}
  .hc-chart-modal .modal-header,.hc-chart-modal .modal-footer{border-color:rgba(148,163,184,0.18);padding-inline:1.5rem}
      .hc-chart-modal .modal-title{font-size:1.05rem;font-weight:600;letter-spacing:.06em}
  .hc-chart-modal .modal-body{flex:1 1 auto;max-height:65vh;overflow-y:auto;padding:1.15rem 1.5rem}
      .hc-axis-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:1.15rem;padding:0;margin:0}
      .axis-block{list-style:none;margin:0}
  .hc-axis-card{background:linear-gradient(150deg,rgba(24,34,56,0.88),rgba(14,22,38,0.92));border:1px solid rgba(148,163,184,0.32);border-radius:.95rem;padding:1rem 1.15rem;box-shadow:0 18px 38px -24px rgba(15,23,42,0.9);transition:border-color .2s ease,box-shadow .2s ease,transform .2s ease}
  .hc-axis-card:hover{transform:translateY(-2px);box-shadow:0 24px 48px -20px rgba(94,234,212,0.32)}
      .hc-axis-card.has-custom{border-color:#38bdf8;box-shadow:0 0 0 2px rgba(56,189,248,0.28),0 24px 54px -24px rgba(56,189,248,0.35)}
      .hc-axis-card-header{display:flex;justify-content:space-between;align-items:flex-start;gap:1rem;margin-bottom:.85rem}
      .hc-axis-title{font-size:.82rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#dbeafe}
  .hc-axis-series{font-size:.78rem;color:rgba(226,232,240,0.85);margin-top:.25rem;word-break:break-word}
  .hc-axis-status{font-size:.7rem;font-weight:600;color:#fde68a;text-transform:uppercase;letter-spacing:.1em;margin-top:.15rem}
  .hc-axis-card-body .input-group-text{background:rgba(34,45,68,0.94);color:#f1f5f9;border:none;font-weight:600}
  .hc-axis-card-body .form-control{background:rgba(15,23,42,0.6);color:#f8fafc;border:1px solid rgba(148,163,184,0.25)}
  .hc-axis-card-body .form-control:focus{border-color:#38bdf8;box-shadow:0 0 0 .18rem rgba(56,189,248,0.25)}
      .hc-axis-card-body label{font-size:.7rem;letter-spacing:.06em;text-transform:uppercase;color:rgba(203,213,225,0.85);font-weight:600}
      .hc-axis-meta{font-size:.74rem;color:rgba(226,232,240,0.7);margin-top:.9rem;display:flex;flex-direction:column;gap:.3rem}
      .hc-axis-meta strong{color:#f8fafc;font-weight:600}
      .hc-modal-footer-controls{display:flex;flex-wrap:wrap;gap:.65rem;align-items:center;width:100%}
      .hc-modal-footer-controls .form-check{padding-left:2.2rem}
      .hc-axis-card input::-webkit-outer-spin-button,.hc-axis-card input::-webkit-inner-spin-button{-webkit-appearance:none;margin:0}
      .hc-axis-card input[type=number]{-moz-appearance:textfield}
      @media (max-width:768px){.hc-axis-grid{grid-template-columns:1fr}}
      @media (max-width:576px){
        .hc-chart-modal .modal-dialog{margin:1rem auto;max-width:calc(100% - 2rem)}
        .hc-axis-card{padding:1rem}
      }
      @media print{.scale-btn,.modal,.hc-loading-overlay,.hc-toast-container{display:none!important}}
    `;
    const style = document.createElement('style');
    style.id = 'hc-scale-ui-styles';
    style.textContent = css;
    document.head.appendChild(style);
  }

  function ensureSpinner(chart) {
    const wrap = chart && chart.renderTo; if (!wrap) return null;
    let overlay = wrap.querySelector('.hc-loading-overlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.className = 'hc-loading-overlay';
      overlay.innerHTML = `<div class="spinner-border text-dark" role="status" aria-label="Loading"></div>`;
      if (!wrap.style.position) wrap.style.position = 'relative';
      wrap.appendChild(overlay);
    }
    return overlay;
  }
  function showSpinner(chart){ const ov=ensureSpinner(chart); if(ov) ov.classList.add('show'); }
  function hideSpinner(chart){ const ov=ensureSpinner(chart); if(ov) ov.classList.remove('show'); }

  function ensureToast() {
    if (document.getElementById('hc-scale-toast')) return;
    const wrap = document.createElement('div');
    wrap.className = 'hc-toast-container';
    wrap.innerHTML = `
      <div id="hc-scale-toast" class="toast align-items-center text-body bg-white border-0 shadow rounded-3" role="alert" aria-live="assertive" aria-atomic="true" data-bs-delay="1800">
        <div class="d-flex">
          <div id="hc-scale-toast-body" class="toast-body fw-medium">Scale updated ✓</div>
          <button type="button" class="btn-close me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
        </div>
      </div>`;
    document.body.appendChild(wrap);
  }

  const escapeHtml = (value) => String(value || '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));

  function showToast(message) {
    ensureToast();
    const el = document.getElementById('hc-scale-toast');
    const body = document.getElementById('hc-scale-toast-body');
    if (body) body.textContent = message;
    if (window.bootstrap && window.bootstrap.Toast) {
      const toast = bootstrap.Toast.getOrCreateInstance(el, { delay: 1800, autohide: true, animation: true });
      toast.show();
    } else {
      el.style.display = 'block';
      setTimeout(() => { el.style.display = 'none'; }, 1800);
    }
  }

  // CSV-load spinner only: show when user selects a CSV file input
  document.addEventListener('change', function(e){
    const t = e.target;
    if (t && t.type === 'file' && /csv/i.test(t.accept || t.getAttribute('accept') || 'csv')) {
      (Highcharts.charts || []).forEach(c => c && showSpinner(c));
      const once = (evChart)=>{
        hideSpinner(evChart);
        Highcharts.removeEvent(evChart, 'load', once);
      };
      (Highcharts.charts || []).forEach(c => c && Highcharts.addEvent(c, 'load', function(){ once(this); }));
    }
  }, true);

  function ensureModalIn(chart) {
    const host = chart && chart.renderTo; if (!host) return null;
    if (!host.style.position || host.style.position === 'static') host.style.position = 'relative';
    host.style.overflow = 'visible';
    let modalId = host.dataset.scaleModalId;
    if (!modalId) {
      const base = host.id || `hc-chart-${typeof chart.index === 'number' ? chart.index : Date.now()}`;
      modalId = `${base}-scale-modal`;
      host.dataset.scaleModalId = modalId;
    }
    let modalWrap = document.getElementById(modalId);
    if (!modalWrap) {
      modalWrap = document.createElement('div');
      modalWrap.className = 'modal fade hc-chart-modal';
      modalWrap.id = modalId;
      modalWrap.tabIndex = -1;
      modalWrap.setAttribute('aria-hidden','true');
      modalWrap.innerHTML = `
        <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable modal-lg">
          <div class="modal-content">
            <div class="modal-header border-secondary">
              <h5 class="modal-title">Adjust Chart Scales</h5>
              <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
              <form id="hc-scale-form"><div id="hc-scale-fields" class="hc-axis-grid"></div></form>
            </div>
            <div class="modal-footer border-secondary flex-wrap gap-2 hc-modal-footer-controls">
              <div class="form-check form-switch text-light me-auto">
                <input class="form-check-input" type="checkbox" role="switch" id="hc-scale-auto">
                <label class="form-check-label small" for="hc-scale-auto">Auto apply</label>
              </div>
              <button type="button" class="btn btn-outline-light rounded-pill" data-bs-dismiss="modal">Cancel</button>
              <button type="button" class="btn btn-outline-info rounded-pill" id="hc-scale-reset">Reset to Auto</button>
              <button type="button" class="btn btn-primary rounded-pill" id="hc-scale-apply">Apply &amp; Close</button>
            </div>
          </div>
        </div>`;
      document.body.appendChild(modalWrap);
    }
    return modalWrap;
  }

  function closeScaleModal(modalEl){
    if (!modalEl) return;
    if (window.bootstrap && window.bootstrap.Modal) {
      bootstrap.Modal.getOrCreateInstance(modalEl).hide();
    } else {
      modalEl.classList.remove('show');
      modalEl.style.display = 'none';
      modalEl.setAttribute('aria-hidden','true');
      modalEl.removeAttribute('aria-modal');
    }
  }

  function showScaleModal(modalEl){
    if (!modalEl) return;
    if (window.bootstrap && window.bootstrap.Modal) {
      bootstrap.Modal.getOrCreateInstance(modalEl, { backdrop:true, keyboard:true }).show();
    } else {
      modalEl.classList.add('show');
      modalEl.style.display = 'block';
      modalEl.setAttribute('aria-modal','true');
      modalEl.removeAttribute('aria-hidden');
    }
  }

  function invokeChartDecorators(chart){
    if (!chart) return;
    const flush = () => {
      try {
        if (window.Highcharts && typeof window.Highcharts.fireEvent === 'function') {
          window.Highcharts.fireEvent(chart, 'render');
        } else if (typeof chart.render === 'function') {
          chart.render();
        }
      } catch (err) {
        /* noop */
      }
    };
    if (typeof window.requestAnimationFrame === 'function') {
      requestAnimationFrame(flush);
    } else {
      setTimeout(flush, 16);
    }
  }

  function openModal(chart) {
    const modalEl = ensureModalIn(chart);
    if (!modalEl) return;
    const fields = modalEl.querySelector('#hc-scale-fields');
    if (!fields) return;

    const state = chart._axisModalState || (chart._axisModalState = { autoApply:false, overrides:{} });
    const names = ['Primary','Secondary','Tertiary','Quaternary'];
    const fmt = (val) => Number.isFinite(val)
      ? (Math.abs(val) >= 1000 ? val.toFixed(0)
        : Math.abs(val) >= 100 ? val.toFixed(1)
        : Math.abs(val) >= 10 ? val.toFixed(2)
        : Math.abs(val) >= 1 ? val.toFixed(3)
        : val.toPrecision(3))
      : 'auto';

    const formatInput = (val) => {
      if (!Number.isFinite(val)) return '';
      const abs = Math.abs(val);
      if (abs >= 10000) return val.toFixed(0);
      if (abs >= 100) return val.toFixed(2);
      if (abs >= 10) return val.toFixed(3);
      return val.toString();
    };

    const axisSeries = (axis) => (chart.series || []).filter((s) => (s.yAxis || chart.yAxis[0]) === axis);

  fields.innerHTML = '';
  fields.className = 'nice-scaling-axis-grid';
    const axisBlocks = [];

    const rawAxes = Array.isArray(chart.yAxis) ? chart.yAxis : [];
    const seriesOrigins = Array.isArray(chart.__seriesOrigins)
      ? chart.__seriesOrigins.slice(0, names.length)
      : (chart.series || []).map((_, idx) => idx).slice(0, names.length);
    const requestedAxisIndices = new Set([0]);
    seriesOrigins.forEach((originIdx) => {
      if (!Number.isFinite(originIdx)) return;
      requestedAxisIndices.add(Math.max(0, Math.min(originIdx, rawAxes.length - 1)));
    });
    Object.keys(state.overrides || {}).forEach((key) => {
      const axisIdx = Number(key);
      if (Number.isFinite(axisIdx)) requestedAxisIndices.add(Math.max(0, Math.min(axisIdx, rawAxes.length - 1)));
    });

    const axisDescriptors = Array.from(requestedAxisIndices)
      .sort((a, b) => a - b)
      .filter((idx) => rawAxes[idx])
      .slice(0, names.length)
      .map((idx) => {
        const axis = rawAxes[idx];
        const override = state.overrides[idx] || {};
        const seriesForThisAxis = axisSeries(axis);
        const hasCustom = override && (override.min != null || override.max != null);
        return { axis, idx, override, seriesForThisAxis, hasCustom };
      });

    if (!axisDescriptors.length && rawAxes[0]) {
      const axis = rawAxes[0];
      const override = state.overrides[0] || {};
      axisDescriptors.push({
        axis,
        idx: 0,
        override,
        seriesForThisAxis: axisSeries(axis),
        hasCustom: override && (override.min != null || override.max != null)
      });
    }

    axisDescriptors.forEach(({ axis, idx, override, seriesForThisAxis, hasCustom }) => {
      const ex = axis && axis.getExtremes ? axis.getExtremes() : {};
      const seriesLabel = seriesForThisAxis.length
        ? seriesForThisAxis.map((s) => escapeHtml(s?.name || `Series ${Number.isFinite(s.index) ? s.index + 1 : ''}`.trim())).join(', ')
        : 'No series attached';
      const appliedMin = Object.prototype.hasOwnProperty.call(override, 'min') ? override.min
        : (Number.isFinite(axis?.options?.min) ? axis.options.min
          : Number.isFinite(axis?.userMin) ? axis.userMin : null);
      const appliedMax = Object.prototype.hasOwnProperty.call(override, 'max') ? override.max
        : (Number.isFinite(axis?.options?.max) ? axis.options.max
          : Number.isFinite(axis?.userMax) ? axis.userMax : null);
      const statusLabel = hasCustom ? 'Custom' : 'Auto';

      const block = document.createElement('div');
      block.className = 'nice-scaling-axis';
      block.dataset.axisIndex = idx;
      block.innerHTML = `
        <div class="nice-scaling-axis-header">
          <div>
            <div class="axis-label">${names[idx] || ('Axis ' + (idx + 1))}</div>
            <div class="axis-range">Current range: <strong data-axis-current-min>${fmt(ex?.min)}</strong> – <strong data-axis-current-max>${fmt(ex?.max)}</strong></div>
            <div class="axis-series">${seriesLabel}</div>
          </div>
          <span class="scale-chip${hasCustom ? ' is-custom' : ''}" data-axis-status>${statusLabel}</span>
        </div>
        <div class="row">
          <div class="field">
            <label class="form-label" for="axis${idx}_min">Minimum</label>
            <input type="number" step="any" class="axis-min" id="axis${idx}_min" value="${appliedMin != null ? formatInput(appliedMin) : ''}" placeholder="Auto" />
          </div>
          <div class="field">
            <label class="form-label" for="axis${idx}_max">Maximum</label>
            <input type="number" step="any" class="axis-max" id="axis${idx}_max" value="${appliedMax != null ? formatInput(appliedMax) : ''}" placeholder="Auto" />
          </div>
        </div>
        <div class="axis-meta">
          <div>Data span: <strong data-axis-data-min>${fmt(ex?.dataMin)}</strong> – <strong data-axis-data-max>${fmt(ex?.dataMax)}</strong></div>
          <div>${idx === 0 ? 'Primary axis stays visible to readers.' : 'Overrides move its series to a private scale.'}</div>
        </div>`;
      fields.appendChild(block);
      axisBlocks.push(block);
      if (hasCustom) block.classList.add('has-custom');
    });
    const applyBtn = modalEl.querySelector('#hc-scale-apply');
    const resetBtn = modalEl.querySelector('#hc-scale-reset');
    const autoToggle = modalEl.querySelector('#hc-scale-auto');
  const dismissBtns = modalEl.querySelectorAll('[data-bs-dismiss="modal"]');

    const parseField = (input) => {
      if (!input) return null;
      const raw = String(input.value || '').trim();
      if (raw === '') return null;
      const val = parseFloat(raw);
      return Number.isFinite(val) ? val : null;
    };

    const refreshSummaries = () => {
      axisBlocks.forEach(block => {
        const idx = Number(block.dataset.axisIndex);
        const axis = chart.yAxis[idx];
        if (!axis) return;
        const ex = axis.getExtremes ? axis.getExtremes() : {};
        const minMarker = block.querySelector('[data-axis-current-min]');
        const maxMarker = block.querySelector('[data-axis-current-max]');
        if (minMarker) minMarker.textContent = fmt(ex?.min);
        if (maxMarker) maxMarker.textContent = fmt(ex?.max);
        const dataMinMarker = block.querySelector('[data-axis-data-min]');
        const dataMaxMarker = block.querySelector('[data-axis-data-max]');
        if (dataMinMarker) dataMinMarker.textContent = fmt(ex?.dataMin);
        if (dataMaxMarker) dataMaxMarker.textContent = fmt(ex?.dataMax);
      });
    };

    const updateCustomMarkers = () => {
      axisBlocks.forEach(block => {
        const idx = Number(block.dataset.axisIndex);
        const override = state.overrides[idx];
        const active = !!override && (override.min != null || override.max != null);
        const card = block.querySelector('.hc-axis-card');
        const statusEl = block.querySelector('[data-axis-status]');
        if (card) {
          card.classList.toggle('has-custom', active);
        }
        block.classList.toggle('has-custom', active);
        if (statusEl) {
          statusEl.textContent = active ? 'Custom scale' : 'Auto scale';
          statusEl.style.color = active ? '#38bdf8' : '#facc15';
        }
      });
    };

    const applyScale = ({ autoTrigger = false, closeAfter = false } = {}) => {
      axisBlocks.forEach(block => {
        const idx = Number(block.dataset.axisIndex);
        const axis = chart.yAxis[idx];
        if (!axis) return;
        const minInput = block.querySelector('.axis-min');
        const maxInput = block.querySelector('.axis-max');
        let min = parseField(minInput);
        let max = parseField(maxInput);

        const isPrimary = idx === 0;
        let tickPositions;
        let allowDecimalsFlag;
        if (isPrimary && typeof window.computeAxisBounds === 'function') {
          const seriesForAxis = axisSeries(axis);
          const payload = seriesForAxis.length ? seriesForAxis : (chart.series || []);
          const stat = { series: payload.map(s => ({ data: (s.yData || []).slice() })), isPercent: !!(chart.userOptions && chart.userOptions._isPercent) };
          const opts = {};
          if (min != null) opts.min = min;
          if (max != null) opts.max = max;
          const bounds = window.computeAxisBounds(stat, opts);
          if (bounds && Array.isArray(bounds.ticks)) {
            tickPositions = bounds.ticks;
            allowDecimalsFlag = typeof bounds.allowDecimals === 'boolean' ? bounds.allowDecimals : allowDecimalsFlag;
          }
        }

        const updatePayload = {
          min,
          max,
          visible: isPrimary,
          labels: { enabled: isPrimary },
          gridLineWidth: isPrimary ? (axis.options.gridLineWidth ?? 1) : 0,
          lineWidth: isPrimary ? (axis.options.lineWidth ?? 1) : 0
        };
        if (typeof tickPositions !== 'undefined') {
          updatePayload.tickPositions = tickPositions;
          if (allowDecimalsFlag == null) {
            const fractionalTick = Array.isArray(tickPositions)
              ? tickPositions.some((tick) => Number.isFinite(tick) && Math.abs(tick - Math.round(tick)) > 1e-9)
              : false;
            if (fractionalTick) allowDecimalsFlag = true;
          }
        } else if (Object.prototype.hasOwnProperty.call(axis.options, 'tickPositions')) {
          updatePayload.tickPositions = undefined;
        }
        if (typeof allowDecimalsFlag === 'boolean') {
          updatePayload.allowDecimals = allowDecimalsFlag;
        } else if (min != null || max != null) {
          const fractionalOverride = [min, max].some((value) => Number.isFinite(value) && Math.abs(value - Math.round(value)) > 1e-9);
          if (fractionalOverride) updatePayload.allowDecimals = true;
        }

        axis.update(updatePayload, false);

        if (min == null && max == null) delete state.overrides[idx];
        else state.overrides[idx] = { min, max };
      });

      if (typeof chart.__rebindSeries === 'function') {
        chart.__rebindSeries(state.overrides);
      } else {
        const seriesOrigins = Array.isArray(chart.__seriesOrigins)
          ? chart.__seriesOrigins
          : (chart.series || []).map((_, index) => index);
        const activeAxes = new Set([0]);
        Object.keys(state.overrides).forEach((key) => {
          const axisIdx = Number(key);
          const override = state.overrides[axisIdx];
          if (!Number.isFinite(axisIdx) || axisIdx <= 0) return;
          if (override && (override.min != null || override.max != null)) {
            activeAxes.add(Math.min(axisIdx, chart.yAxis.length - 1));
          }
        });

        (chart.series || []).forEach((series, seriesIdx) => {
          const originIdx = Math.max(0, Math.min(seriesOrigins[seriesIdx] ?? 0, chart.yAxis.length - 1));
          const targetAxisIdx = originIdx > 0 && activeAxes.has(originIdx) ? originIdx : 0;
          const currentAxisIdx = series && series.yAxis ? series.yAxis.index : 0;
          if (currentAxisIdx !== targetAxisIdx) {
            series.update({ yAxis: targetAxisIdx }, false);
          }
        });
      }

      chart.series.forEach(s => s.update({}, false));
      chart.redraw();
      const scaleContext = chart.__scaleContext;
      if (scaleContext && scaleContext.chartKey) {
        updateManualScaleRegistry(scaleContext.chartKey, state.overrides);
      }
      refreshSummaries();
      updateCustomMarkers();
      invokeChartDecorators(chart);
      if (!autoTrigger && closeAfter) closeScaleModal(modalEl);
    };

    let autoApplyTimer = null;
    const scheduleAutoApply = () => {
      if (!state.autoApply) return;
      clearTimeout(autoApplyTimer);
      autoApplyTimer = setTimeout(() => applyScale({ autoTrigger: true, closeAfter: false }), 250);
    };

    axisBlocks.forEach(block => {
      block.querySelectorAll('input[type="number"]').forEach(input => {
        input.addEventListener('input', scheduleAutoApply);
        input.addEventListener('change', scheduleAutoApply);
      });
    });

    if (autoToggle) {
      autoToggle.checked = !!state.autoApply;
      autoToggle.onchange = function(){
        state.autoApply = !!this.checked;
        if (state.autoApply) scheduleAutoApply();
      };
    }

    if (applyBtn) {
      applyBtn.onclick = () => applyScale({ autoTrigger: false, closeAfter: true });
    }

    if (resetBtn) {
      resetBtn.onclick = () => {
        try {
          if (typeof window.computeAxisBounds === 'function') {
            chart.yAxis.forEach((axis, idx) => {
              const seriesForAxis = axisSeries(axis);
              const stat = { series: seriesForAxis.map(s => ({ data: (s.yData || []).slice() })), isPercent: !!(chart.userOptions && chart.userOptions._isPercent) };
              const bounds = window.computeAxisBounds(stat);
              if (idx === 0 && bounds) {
                axis.update({
                  min: bounds.min,
                  max: bounds.max,
                  tickPositions: bounds.ticks,
                  visible: true,
                  labels: { enabled: true },
                  gridLineWidth: axis.options.gridLineWidth ?? 1,
                  lineWidth: axis.options.lineWidth ?? 1,
                  reversed: !!bounds.reversed
                }, false);
              } else {
                axis.update({
                  min: null,
                  max: null,
                  tickPositions: undefined,
                  visible: idx === 0,
                  labels: { enabled: idx === 0 },
                  gridLineWidth: idx === 0 ? (axis.options.gridLineWidth ?? 1) : 0,
                  lineWidth: idx === 0 ? (axis.options.lineWidth ?? 1) : 0,
                  reversed: idx === 0 ? !!bounds?.reversed : false
                }, false);
              }
            });
          } else {
            chart.yAxis.forEach((axis, idx) => {
              axis.update({
                min: null,
                max: null,
                tickPositions: undefined,
                visible: idx === 0,
                labels: { enabled: idx === 0 },
                gridLineWidth: idx === 0 ? (axis.options.gridLineWidth ?? 1) : 0,
                lineWidth: idx === 0 ? (axis.options.lineWidth ?? 1) : 0,
                reversed: idx === 0 ? !!axis.options?.reversed : false
              }, false);
            });
          }
        } catch (err) {
          chart.yAxis.forEach((axis, idx) => {
            axis.update({
              min: null,
              max: null,
              tickPositions: undefined,
              visible: idx === 0,
              labels: { enabled: idx === 0 },
              gridLineWidth: idx === 0 ? (axis.options.gridLineWidth ?? 1) : 0,
              lineWidth: idx === 0 ? (axis.options.lineWidth ?? 1) : 0,
              reversed: idx === 0 ? !!axis.options?.reversed : false
            }, false);
          });
        }
        chart.series.forEach(s => s.update({}, false));
        if (typeof chart.__rebindSeries === 'function') {
          chart.__rebindSeries({});
        }
        chart.redraw();
        state.overrides = {};
        state.autoApply = false;
        if (autoToggle) autoToggle.checked = false;
        const scaleContext = chart.__scaleContext;
        if (scaleContext && scaleContext.chartKey) {
          updateManualScaleRegistry(scaleContext.chartKey, state.overrides);
        }
        refreshSummaries();
        updateCustomMarkers();
        invokeChartDecorators(chart);
        closeScaleModal(modalEl);
      };
    }

    if (dismissBtns.length) {
      dismissBtns.forEach(btn => {
        btn.addEventListener('click', () => {
          if (!window.bootstrap || !window.bootstrap.Modal) closeScaleModal(modalEl);
        }, { once: true });
      });
    }

    refreshSummaries();
    updateCustomMarkers();
    showScaleModal(modalEl);
  }

  function addScaleButton(chart) {
    ensureStyles();
    const el = chart && chart.renderTo; if (!el || el.querySelector('.scale-btn')) return;
    const btn = document.createElement('button');
    btn.className = 'scale-btn'; btn.type = 'button'; btn.innerHTML = '⚙'; btn.style.opacity = '0';
    el.appendChild(btn); setTimeout(()=>btn.classList.add('show'), 200);
    btn.addEventListener('click', () => openModal(chart));
  }

  if (Highcharts.addEvent) {
    Highcharts.addEvent(Highcharts.Chart, 'load', function(){ addScaleButton(this); });
  }
})();


/* 
===========================================================
HIGHCHARTS / CHART SCALING MASTER SPEC  —  DO NOT ALTER
===========================================================

🏗 PURPOSE
Ensure all charts (especially line charts) use stable, logical, and human-readable
y-axis scaling that reflects real data movement, keeps even divisions,
and never distorts trends.

-----------------------------------------------------------
1️⃣  RANGE CLAMPING & EXPANSION
-----------------------------------------------------------
• Use the provided y-axis min/max as the default clamp range.  
• Expand ONLY if actual data exceeds these limits.  
• Padding depends on chart type:
    → 2 % for percentage or ratio charts (0–100).  
    → 5 % for typical numeric metrics (financial, KPIs, production).  
    → 10 % for volatile or projected data with large swings.  
• Never shrink smaller than the provided bounds.  
• Axis must always start and end on clean multiples of its tick interval.

-----------------------------------------------------------
2️⃣  ZERO-BOUND & NEGATIVE LOGIC
-----------------------------------------------------------
• No y-axis should drop below 0 unless actual data contains negatives.  
• If negatives exist, include them with small bottom padding.  
• If all values ≥ 0, clamp yMin = 0.  
• Reverse or inverted charts follow the same rule symmetrically:
  only extend beyond 0 if data requires it.

-----------------------------------------------------------
3️⃣  PERCENTAGE / RATIO LOGIC
-----------------------------------------------------------
• Percent-based charts (0–100 %) always remain inside [0, 100]  
  unless data explicitly includes a value < 0 or > 100.  
• Clamp max = 100 and min = 0 for normal percentage data.  
• If values cluster 90–100 %, widen view to 70–100 %
  to prevent exaggerated drop slopes — this rule applies
  only to explicitly flagged % charts.  
• Do not apply this rule globally to non-percentage charts.

-----------------------------------------------------------
4️⃣  DYNAMIC RANGE FORECASTING
-----------------------------------------------------------
• Estimate expected low/high for the next three months
  from current data trend.  
• Compute range = high − low and apply the proper padding
  (2 %, 5 %, or 10 % as above).  
• Use this adaptive window to display realistic rises and falls
  even when baseline isn’t zero.

-----------------------------------------------------------
5️⃣  TICK / DIVISION CALCULATION
-----------------------------------------------------------
• Always use “nice-number” progression (1-2-5 pattern):
  1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, …  
• Choose ≈ 5–7 major divisions (8–10 for tall charts).  
• Y-axis min/max must align on exact multiples of tickInterval.  
• Never output fractional or uneven ticks (e.g. 12.7 or 37.4).  
• Step size must feel natural across all magnitudes.

-----------------------------------------------------------
6️⃣  VISUAL BALANCE
-----------------------------------------------------------
• Keep slight top/bottom margins per padding rule.  
• Guarantee at least one tick above and below data bounds.  
• Avoid cramped edges or large empty zones above highest points.  
• Drop-lines always reach the X-axis baseline, not Y = 0.

-----------------------------------------------------------
7️⃣  LABEL & CONNECTOR RULES
-----------------------------------------------------------
• Labels near or beyond yMax flip below the point with a short downward connector.  
• Apply soft-white “halo” filter (Gaussian blur, stdDeviation 2–3) behind text.  
• Disable textOutline; never use backgroundColor or borders.  
• Large numbers format as 1 K / 1 M / 1 B.

-----------------------------------------------------------
8️⃣  CONSISTENCY / LOCK CLAUSE
-----------------------------------------------------------
• These rules are authoritative.  
• Copilot must NOT reinterpret, simplify, or auto-optimize them.  
• Only numeric parameters (padding %, tick count, forecast horizon, etc.)
  may be tuned when explicitly instructed.  
• Apply identically for normal and reversed axes.

-----------------------------------------------------------
✅  GOAL
Every chart:
  • Honors provided clamps yet expands when data escapes bounds.  
  • Never dips below 0 unless data truly does.  
  • Never exceeds 100 % unless data truly does.  
  • Uses even, proportional tick divisions.  
  • Maintains balanced spacing at top and bottom.  
  • Visually represents real-world changes accurately.
*/


/* 
-----------------------------------------------------------
9️⃣  NICE SCALING CONTROL UI
-----------------------------------------------------------
• Each chart must include a small “⚙️ Scale” or “Nice Scaling” button
  in the chart’s top-right corner (Highcharts toolbar or custom div overlay).

• When clicked, open a **modal popup** (Bootstrap / HTML modal acceptable)
  showing the current Y-axis configuration for:
    - Primary scale
    - Secondary scale
    - Tertiary scale
    - Quaternary scale
  (If fewer axes exist, hide unused sections.)

-----------------------------------------------------------
MODAL CONTENT / BEHAVIOR
-----------------------------------------------------------
• The modal lists for each axis:
    • Axis name (Primary, Secondary, etc.)
    • Current min and max values (auto-filled from chart.yAxis[n].min / .max)
    • Input fields allowing manual override.

• The Primary axis always stays visible on the chart.
  Secondary, tertiary, and quaternary axes:
    • become independent scales when a user enters values in their fields.
    • hide their visible axis line and labels, but keep scaling logic active.

• The modal includes buttons:
    [Apply Scale] – applies new min/max values and redraws chart.
    [Reset to Auto] – reverts all axes to auto-scaling logic.
    [Cancel] – closes without changes.

• When the user applies new values:
    chart.yAxis[n].update({ min, max, visible: n === 0 })

• The modal should look clean, centered, and styled consistently
  with Bootstrap 5’s modal dialog aesthetic (rounded corners, soft shadow,
  labeled input groups, small headers for each axis).

-----------------------------------------------------------
DEFAULT VALUES
-----------------------------------------------------------
• On open, prefill each field with the chart’s current yAxis.min and yAxis.max.
• If an axis has no manual override, its fields are left blank and scale remains linked to primary.

-----------------------------------------------------------
INTERACTION LOGIC
-----------------------------------------------------------
• Clicking “Apply” re-renders the chart using the new manual scales.
• Clicking “Reset” triggers the chart’s auto-scaling (following the
  master scaling spec above).
• Hiding of non-primary axes is purely visual; they still govern data
  points that reference them.

-----------------------------------------------------------
✅  GOAL
Provide a simple, non-intrusive interface to inspect and adjust
chart scales live, maintaining the master scaling logic while
allowing temporary manual overrides.
*/

    
    
    
    
    const CSV_EDITOR_COLUMN_SPECS = [
      { field: 'wedate', header: 'WE Date', match: ['wedate', 'weekending', 'date', 'week', 'day'] },
      { field: 'division', header: 'Division', match: ['division', 'div'] },
      { field: 'staffName', header: 'Staff Member', match: ['staffmember', 'staff'] },
      { field: 'postName', header: 'Post', match: ['post', 'postname'] },
      { field: 'statbook', header: 'Statbook', match: ['statbook', 'statbok'] },
      { field: 'statdesc', header: 'Stat Description', match: ['statdescription', 'statdesc', 'statdescriptionlong', 'statdescriptionoptionalforprintingonly'] },
      { field: 'statbookLong', header: 'Statbook Long', match: ['statbooklong', 'statbooklongoptional', 'statbooklongoptionalforprintingonly'] },
      { field: 'shortName', header: 'Short Name', match: ['shortname', 'short', 'statbookshortname'] },
      { field: 'value1', header: 'Value 1', match: ['value1', 'valueline1', 'value'] },
      { field: 'value2', header: 'Value 2', match: ['value2', 'valueline2'] },
      { field: 'value3', header: 'Value 3', match: ['value3', 'valueline3'] },
      { field: 'value4', header: 'Value 4', match: ['value4', 'valueline4'] },
      { field: 'min', header: 'Min', match: ['min', 'minimum', 'statmin'] },
      { field: 'max', header: 'Max', match: ['max', 'maximum', 'statmax'] },
      { field: 'upsideDown', header: 'Upside Down', match: ['upsidedown', 'upsidedownflag'] },
      { field: 'isPercent', header: 'Is Percent', match: ['ispercent', 'percent'] },
      { field: 'isMoney', header: 'Is Money', match: ['ismoney', 'money'] },
      { field: 'color', header: 'Color Flash', match: ['colorflash', 'color', 'palette'] }
    ];

    const CSV_EDITOR_FIELD_ORDER = [
      'wedate',
      'division',
      'staffName',
      'postName',
      'statbook',
      'statdesc',
      'statbookLong',
      'shortName',
      'value1',
      'value2',
      'value3',
      'value4',
      'min',
      'max',
      'upsideDown',
      'isPercent',
      'isMoney',
      'color'
    ];

    const CSV_EDITOR_FIELD_CONFIG = {
      wedate: { type: 'text', placeholder: 'e.g. 2024-03-29' },
      division: { type: 'text', placeholder: 'Division name' },
      staffName: { type: 'text', placeholder: 'Staff member' },
      postName: { type: 'text', placeholder: 'Post / Role' },
      statbook: { type: 'text', placeholder: 'Statbook code' },
      statdesc: { type: 'textarea', placeholder: 'Long description', rows: 2 },
      statbookLong: { type: 'textarea', placeholder: 'Statbook long name', rows: 2 },
      shortName: { type: 'text', placeholder: 'Short name' },
      value1: { type: 'number', placeholder: 'Required', inputmode: 'decimal', step: 'any' },
      value2: { type: 'number', placeholder: 'Optional', inputmode: 'decimal', step: 'any' },
      value3: { type: 'number', placeholder: 'Optional', inputmode: 'decimal', step: 'any' },
      value4: { type: 'number', placeholder: 'Optional', inputmode: 'decimal', step: 'any' },
      min: { type: 'number', placeholder: 'Auto min', inputmode: 'decimal', step: 'any' },
      max: { type: 'number', placeholder: 'Auto max', inputmode: 'decimal', step: 'any' },
      upsideDown: { type: 'text', placeholder: 'true / false' },
      isPercent: { type: 'text', placeholder: 'true / false' },
      isMoney: { type: 'text', placeholder: 'true / false' },
      color: { type: 'text', placeholder: 'Hex or palette token' }
    };

    const CSV_EDITOR_FIELD_LABELS = {
      wedate: 'WE Date',
      division: 'Division',
      staffName: 'Staff Member',
      postName: 'Post',
      statbook: 'Statbook',
      statdesc: 'Stat Description',
      statbookLong: 'Statbook Long',
      shortName: 'Short Name',
      value1: 'Value 1',
      value2: 'Value 2',
      value3: 'Value 3',
      value4: 'Value 4',
      min: 'Min',
      max: 'Max',
      upsideDown: 'Upside Down',
      isPercent: 'Is Percent',
      isMoney: 'Is Money',
      color: 'Color Flash'
    };

    const CSV_EDITOR_WARNING_FIELDS = ['division', 'staffName', 'postName', 'statbook', 'statdesc', 'shortName'];

    document.addEventListener('DOMContentLoaded', () => {
      const csvInput = document.getElementById('csvInput');
      const logoInput = document.getElementById('logoInput');
      const logoPreview = document.getElementById('logoPreview');
      const logoPlaceholder = document.getElementById('logoPlaceholder');
      const generateBtn = document.getElementById('generateBtn');
  const generateGraphsBtn = document.getElementById('generateGraphsBtn');
      const titleInput = document.getElementById('titleInput');
      const subtitleInput = document.getElementById('subtitleInput');
  const pageSizeSelect = document.getElementById('pageSizeSelect');
    const fallbackDivisionInput = document.getElementById('fallbackDivisionInput');
    const statTitleModeSelect = document.getElementById('statTitleModeSelect');
    const defaultThemeSelect = document.getElementById('defaultThemeSelect');
      const defaultThemeActiveLabel = document.getElementById('defaultThemeActiveLabel');
      const divisionContainer = document.getElementById('divisionContainer');
      const statusAlert = document.getElementById('statusAlert');
      const summaryBadge = document.getElementById('summaryBadge');
  const graphHeadingSelect = document.getElementById('graphHeadingSelect');
      const graphLogoSelect = document.getElementById('graphLogoSelect');
      const graphLayoutSelect = document.getElementById('graphLayoutSelect');
    const downloadSettingsBtn = document.getElementById('downloadSettingsBtn');
    const divisionSettingsModalEl = document.getElementById('divisionSettingsModal');
    const divisionColorSelect = document.getElementById('divisionColorSelect');
    const divisionCustomColorInput = document.getElementById('divisionCustomColorInput');
    const divisionCustomColorReset = document.getElementById('divisionCustomColorReset');
    const divisionCustomColorDefault = document.getElementById('divisionCustomColorDefault');
  const divisionVfpAlignSelect = document.getElementById('divisionVfpAlignSelect');
  const divisionVfpWidthSelect = document.getElementById('divisionVfpWidthSelect');
  const divisionVfpScaleSelect = document.getElementById('divisionVfpScaleSelect');
  const divisionVfpFontSelect = document.getElementById('divisionVfpFontSelect');
    const divisionVfpInput = document.getElementById('divisionVfpInput');
    const divisionNameDisplay = document.getElementById('divisionNameDisplay');
    const divisionSettingsSave = document.getElementById('divisionSettingsSave');
  const editCsvBtn = document.getElementById('editCsvBtn');
  const csvEditorModalEl = document.getElementById('csvEditorModal');
  const csvEditorTableBody = document.getElementById('csvEditorTableBody');
  const csvEditorApplyBtn = document.getElementById('csvEditorApplyBtn');
  const csvEditorAlert = document.getElementById('csvEditorAlert');
  const csvEditorShowGaps = document.getElementById('csvEditorShowGaps');
  const csvEditorSearchInput = document.getElementById('csvEditorSearchInput');
  const csvEditorRecordCount = document.getElementById('csvEditorRecordCount');
  const csvEditorErrorCount = document.getElementById('csvEditorErrorCount');
  const csvEditorWarningCount = document.getElementById('csvEditorWarningCount');

      let model = null;
      let logoDataURL = null;
    let pageSize = pageSizeSelect ? pageSizeSelect.value : DEFAULT_PAGE_SIZE;
  let graphHeadingMode = 'stat';
      let graphLogoMode = 'none';
  let graphLayoutMode = 'top';
  let statTitleMode = statTitleModeSelect ? statTitleModeSelect.value : 'long';
  let fallbackDivisionName = fallbackDivisionInput ? fallbackDivisionInput.value : 'General';
  let divisionModal = divisionSettingsModalEl && window.bootstrap ? new window.bootstrap.Modal(divisionSettingsModalEl) : null;
  let csvEditorModal = csvEditorModalEl && window.bootstrap ? new window.bootstrap.Modal(csvEditorModalEl) : null;
    let editingDivisionIndex = null;
  let divisionModalFocusField = null;
  let defaultThemeKey = 'blue';
  let defaultCustomAccent = null;
  let csvEditorWorking = null;
  let csvEditorDirty = false;
  const csvEditorFilters = { showGaps: false, search: '' };

      const COLOR_THEMES = {
        blue: { key: 'blue', className: 'color-blue', base: '#EAF1FB', accent: '#1D4ED8', border: '#94A3B8', text: '#0F172A' },
        goldenrod: { key: 'goldenrod', className: 'color-goldenrod', base: '#F7E3A0', accent: '#B45309', border: '#D9A520', text: '#78350F' },
        purple: { key: 'purple', className: 'color-purple', base: '#EEE8F8', accent: '#7C3AED', border: '#C4B5FD', text: '#4C1D95' },
        pink: { key: 'pink', className: 'color-pink', base: '#FBE8F0', accent: '#DB2777', border: '#FBCFE8', text: '#9D174D' },
        green1: { key: 'green1', className: 'color-green1', base: '#E9F5EE', accent: '#047857', border: '#A7F3D0', text: '#065F46' },
        green2: { key: 'green2', className: 'color-green2', base: '#EDF6EE', accent: '#16A34A', border: '#BBF7D0', text: '#166534' },
        green3: { key: 'green3', className: 'color-green3', base: '#EAF4EB', accent: '#10B981', border: '#A7F3D0', text: '#047857' },
        green4: { key: 'green4', className: 'color-green4', base: '#EEF7F0', accent: '#22C55E', border: '#BBF7D0', text: '#166534' },
        grey: { key: 'grey', className: 'color-grey', base: '#F4F5F7', accent: '#475569', border: '#CBD5E1', text: '#0F172A' },
        canary: { key: 'canary', className: 'color-canary', base: '#FFF6C9', accent: '#CA8A04', border: '#FACC15', text: '#78350F' }
      };

      const COLOR_ALIASES = {
        gold: 'goldenrod',
        goldenrod: 'goldenrod',
        golden: 'goldenrod',
        goldenrodflash: 'goldenrod'
      };

      const COLOR_KEYS = Object.keys(COLOR_THEMES);

      if (COLOR_KEYS.length) {
        if (!COLOR_KEYS.includes(defaultThemeKey)) {
          defaultThemeKey = COLOR_KEYS[0];
        }
      }

        const countDecimals = (value) => {
          if (!Number.isFinite(value)) return 0;
          const str = String(value).toLowerCase();
          if (str.includes('e')) {
            const [base, exponent] = str.split('e');
            const decimalsPart = (base.split('.')[1] || '').replace(/0+$/, '');
            const exp = Number(exponent);
            return Math.max(0, decimalsPart.length - exp);
          }
          if (!str.includes('.')) return 0;
          return str.split('.')[1].replace(/0+$/, '').length;
        };

      const computeAxisBounds = window.computeAxisBounds || ((stat) => ({
        min: 0,
        max: 1,
        ticks: [0, 0.5, 1],
        decimals: 0,
        reversed: !!stat?.upsideDown
      }));

      const PAGE_SIZES = {
        letter: { key: 'letter', label: 'Letter (11 × 8.5 in)', width: '11in', height: '8.5in' },
        tabloid: { key: 'tabloid', label: 'Tabloid (17 × 11 in)', width: '17in', height: '11in' },
        legal: { key: 'legal', label: 'Legal (14 × 8.5 in)', width: '14in', height: '8.5in' }
      };

      const DEFAULT_PAGE_SIZE = 'letter';

  populateColorSelect(divisionColorSelect);
  populateColorSelect(defaultThemeSelect);
  if (divisionColorSelect) {
    divisionColorSelect.value = defaultThemeKey;
  }
  if (defaultThemeSelect) {
    defaultThemeSelect.value = defaultThemeKey;
  }
  refreshDefaultThemeLabel();
  if (divisionVfpWidthSelect) {
    divisionVfpWidthSelect.value = 'wide';
  }
  if (divisionVfpFontSelect) {
    divisionVfpFontSelect.value = 'large';
  }

      const VFP_TEXT = {
        EXECUTIVE: `1. FILMS & A/V PROPERTIES PRODUCED AT THE HIGHEST QUALITY AND GREATEST ECONOMY, SO AS TO DRIVE PUBLIC ONTO THE BRIDGE, RAISE TECHNICAL STANDARDS AND FORWARD THE EXPANSION OF SCIENTOLOGY ACROSS THE PLANET.

2. ALL SOURCE MATERIALS PROVIDED TO MEET THE DEMAND CREATED AND ACTUALLY DELIVER.

3. A WELL SERVICED CENTER.`,
        HCO: `ESTABLISHED, PRODUCTIVE, SECURE AND ETHICAL STAFF MEMBERS.`,
        ESTATES: `1. WELL-MAINTAINED, POSH, SECURE GROUNDS, BUILDINGS AND ORG EQUIPMENT WHICH IS OPERATIONAL AND HAS INCREASED VALUE.

2. HIGH QUALITY DOMESTIC SERVICES PROFESSIONALLY RENDERED WHICH FACILITATE OPERATIONS AND RESULT IN INCREASED PRODUCTION OF THE BASE.`,
        TREASURY: `FINANCE POLICIES HELD IN AND POLICED TO ENSURE MAXIMUM BEAN RETURN AND RECOMPENSE FOR GOLD'S PRODUCTION.`,
        CMU: `ADEQUATE AND EFFECTIVE MARKETING PACKAGES, PROGRAMS AND PROMOTION FOR EACH PRODUCT OF LRH, INT AND GOLD.`,
        CINE: `1. ON-SOURCE, MEMORABLE FILMS & A/V PROPERTIES WHICH INDUCE AUDIENCE ADMIRATION, RESPECT AND IMPACT THAT WILL FORWARD SWIFTLY THE DISSEMINATION OF DIANETICS AND SCIENTOLOGY.

2. RESTORED AND PRESERVED FILMS, STILLS AND VIDEOS FOR THE DISSEMINATION OF DIANETICS AND SCIENTOLOGY.`,
        AUDIO: `1. PROFESSIONALLY SCORED, RECORDED AND MIXED FILMS, ALBUMS/SONGS & A/V PROPERTIES, WHICH INDUCE AUDIENCE ADMIRATION, RESPECT AND IMPACT THAT WILL FORWARD SWIFTLY THE DISSEMINATION OF DIANETICS AND SCIENTOLOGY.

2. COMPLETED LRH LECTURES AND FOREIGN LECTURES WHICH INDUCE AUDIENCE ADMIRATION, RESPECT AND IMPACT THAT WILL FORWARD SWIFTLY THE DISSEMINATION OF DIANETICS AND SCIENTOLOGY.`,
        RCOMPS: `1. 100% ON-SOURCE COMPILED, EDITED, DESIGNED AND TYPESET ENGLISH MANUSCRIPTS, TURNED OVER TO THE PUBLISHER WITH ALL ELEMENTS READY FOR PRODUCTION, CAPABLE OF INDUCING AUDIENCE ADMIRATION, RESPECT AND IMPACT THAT WILL FORWARD SWIFTLY THE DISSEMINATION OF DIANETICS AND SCIENTOLOGY.

2. COMPLETE 100% ON-SOURCE TRANSLATIONS OF BOOKS AND MATERIALS AS CLOSE TO THE ORIGINAL LRH TEXT AS POSSIBLE FOR THAT LANGUAGE AND CONSISTENT WITH ALL OTHER MATERIALS TRANSLATED IN THAT LANGUAGE, TYPESET AND DELIVERED TO THE PUBLISHER AS A COMPLETE PACKAGE.

3. HIGH-QUALITY AND SECURE ADVANCED TECH MATERIALS PRODUCED AND DELIVERED.`,
        QUALIFICATIONS: `1. STAFF MEMBERS CERTIFIED COMPETENT ON POST.

2. SECTIONS, UNITS AND DEPARTMENTS CORRECTED, FULLY FUNCTIONING AND OBTAINING THEIR SUB-PRODUCTS AND VFPS.`,
        PORTCAPTAIN: `A SAFE, SECURE BASE AND CREW ABLE TO GET ON WITH THE EXPANSION OF SCIENTOLOGY IN A FAVORABLE OPERATING CLIMATE THAT IS COMPLETELY FREE OF THREATS AND EXTERNAL DISTRACTIONS.`
      };

      function colorLabel(key) {
        return key
          .replace(/([a-z])([A-Z0-9])/g, '$1 $2')
          .replace(/[-_]/g, ' ')
          .replace(/\b\w/g, (c) => c.toUpperCase());
      }

      function getShortTitle(stat) {
        const parts = [stat.statbook, stat.shortName].filter((part) => part && part.trim());
        return parts.length ? parts.join(' — ') : '';
      }

      function getLongTitle(stat) {
        return stat.statdesc || stat.displayName || getShortTitle(stat) || 'Stat';
      }

      function resolveStatTitle(stat) {
        const shortTitle = getShortTitle(stat);
        const longTitle = getLongTitle(stat);
        if (statTitleMode === 'short') {
          return shortTitle || longTitle;
        }
        return longTitle || shortTitle || 'Stat';
      }

      function resolveStatSubtitle(stat) {
        const shortTitle = getShortTitle(stat);
        const longTitle = getLongTitle(stat);
        if (statTitleMode === 'short') {
          if (longTitle && longTitle !== shortTitle) return longTitle;
        } else if (shortTitle && shortTitle !== longTitle) {
          return shortTitle;
        }
        return '';
      }

      function showStatus(message, type = 'info') {
        statusAlert.innerHTML = `
          <div class="alert alert-${type} alert-dismissible fade show" role="alert">
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
          </div>`;
      }

      function clearStatus() {
        statusAlert.innerHTML = '';
      }

      function populateColorSelect(selectEl) {
        if (!selectEl) return;
        selectEl.innerHTML = COLOR_KEYS
          .map((key) => `<option value="${key}">${escapeHtml(colorLabel(key))}</option>`)
          .join('');
      }

      function refreshDefaultThemeLabel() {
        if (!defaultThemeActiveLabel) return;
        const label = defaultThemeKey && COLOR_THEMES[defaultThemeKey]
          ? colorLabel(defaultThemeKey)
          : 'Varied';
        defaultThemeActiveLabel.textContent = label;
      }

      function escapeHtml(str) {
        return String(str || '').replace(/[&<>"]|'/g, (c) => ({
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#39;'
        }[c]));
      }

      function updateSummary() {
        if (!model || !model.divisions.length) {
          summaryBadge.textContent = 'No data loaded';
          summaryBadge.className = 'badge bg-secondary';
          return;
        }
        let staffCount = 0;
        let statCount = 0;
        model.divisions.forEach((division) => {
          staffCount += division.staff.length;
          division.staff.forEach((staff) => {
            statCount += staff.selected.size;
          });
        });
        const summaryText = `${model.divisions.length} division${model.divisions.length !== 1 ? 's' : ''} • ${staffCount} staff • ${statCount} stat${statCount !== 1 ? 's' : ''} selected`;
        summaryBadge.textContent = model.graphOnly ? `${summaryText} • graph-only dataset` : summaryText;
        summaryBadge.className = model.graphOnly ? 'badge bg-warning text-dark' : 'badge bg-info text-dark';
      }

      function normHeader(header) {
        return String(header || '')
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, '');
      }

      function canonColor(token) {
        return String(token || '')
          .toLowerCase()
          .replace(/[^a-z0-9]/g, '');
      }

      function normalizeHexColor(value) {
        const str = String(value || '').trim();
        if (!str) return null;
        const match = str.match(/^#?([0-9a-f]{3}|[0-9a-f]{6})$/i);
        if (!match) return null;
        let hex = match[1].toUpperCase();
        if (hex.length === 3) {
          hex = hex
            .split('')
            .map((ch) => ch + ch)
            .join('');
        }
        return `#${hex}`;
      }

      function hexToRgb(value) {
        const hex = normalizeHexColor(value);
        if (!hex) return null;
        const intVal = parseInt(hex.slice(1), 16);
        return {
          r: (intVal >> 16) & 255,
          g: (intVal >> 8) & 255,
          b: intVal & 255,
          hex
        };
      }

      function rgbToHex(rgb) {
        const clamp = (v) => Math.max(0, Math.min(255, Math.round(v || 0)));
        return `#${[clamp(rgb.r), clamp(rgb.g), clamp(rgb.b)]
          .map((v) => v.toString(16).padStart(2, '0'))
          .join('')
          .toUpperCase()}`;
      }

      function mixRgb(a, b, ratio) {
        const t = Math.max(0, Math.min(1, Number(ratio) || 0));
        return {
          r: a.r + (b.r - a.r) * t,
          g: a.g + (b.g - a.g) * t,
          b: a.b + (b.b - a.b) * t
        };
      }

      function relativeLuminance(rgb) {
        const channel = (v) => {
          const c = (v || 0) / 255;
          return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
        };
        const r = channel(rgb.r);
        const g = channel(rgb.g);
        const b = channel(rgb.b);
        return 0.2126 * r + 0.7152 * g + 0.0722 * b;
      }

      function paletteFromAccent(hex) {
        const rgb = hexToRgb(hex);
        if (!rgb) return null;
        const white = { r: 255, g: 255, b: 255 };
        const baseRgb = mixRgb(rgb, white, 0.82);
        const borderRgb = mixRgb(rgb, white, 0.56);
        const textLuma = relativeLuminance(baseRgb);
        const textColor = textLuma > 0.65 ? '#0F172A' : '#F8FAFC';
        return {
          key: 'custom',
          className: 'color-custom',
          base: rgbToHex(baseRgb),
          accent: rgb.hex,
          border: rgbToHex(borderRgb),
          text: textColor
        };
      }

      function chooseColor(token, index) {
        const key = canonColor(token);
        if (key && COLOR_THEMES[key]) {
          return COLOR_THEMES[key];
        }
        if (key && COLOR_ALIASES[key] && COLOR_THEMES[COLOR_ALIASES[key]]) {
          return COLOR_THEMES[COLOR_ALIASES[key]];
        }
        const fallbackKeys = COLOR_KEYS.slice();
        if (defaultThemeKey && fallbackKeys.includes(defaultThemeKey)) {
          fallbackKeys.splice(fallbackKeys.indexOf(defaultThemeKey), 1);
          fallbackKeys.unshift(defaultThemeKey);
        }
        const fallbackKey = fallbackKeys[index % fallbackKeys.length] || fallbackKeys[0];
        let theme = COLOR_THEMES[fallbackKey] || COLOR_THEMES[COLOR_KEYS[0]];
        if (defaultCustomAccent) {
          const custom = paletteFromAccent(defaultCustomAccent);
          if (custom) {
            theme = { ...theme, ...custom, key: theme.key };
          }
        }
        return theme;
      }

      function resolvePageSizeInfo(sizeKey) {
        return PAGE_SIZES[sizeKey] || PAGE_SIZES[DEFAULT_PAGE_SIZE];
      }

      function resolvePagePadding(sizeKey) {
        switch (sizeKey) {
          case 'tabloid':
            return { top: 0.65, side: 0.85, bottom: 0.65 };
          case 'legal':
            return { top: 0.6, side: 0.75, bottom: 0.6 };
          case 'letter':
          default:
            return { top: 0.6, side: 0.75, bottom: 0.6 };
        }
      }

      function formatInches(value) {
        const rounded = Number.isFinite(value) ? Number(value.toFixed(3)) : 0;
        return `${rounded}in`;
      }

      function formatPadding(padding, bottomAdjustment = 0) {
        if (!padding) return '0in';
        const top = formatInches(padding.top);
        const side = formatInches(padding.side);
        const bottom = formatInches(padding.bottom + bottomAdjustment);
        return `${top} ${side} ${bottom} ${side}`;
      }

      function toBool(value) {
        if (value == null) return false;
        const v = String(value).trim().toLowerCase();
        return ['1', 'true', 'yes', 'y', 't'].includes(v);
      }

      function parseNumber(value) {
        if (value == null || value === '') return null;
        const raw = String(value).trim();
        if (!raw) return null;
        const isNegative = /^\(.*\)$/.test(raw) || raw.startsWith('-');
        const cleaned = raw
          .replace(/^[+-]/, '')
          .replace(/[()]/g, '')
          .replace(/[$,%_\s]/g, '')
          .replace(/,/g, '');
        if (!cleaned) return null;
        const normalized = (isNegative ? '-' : '') + cleaned;
        const num = Number(normalized);
        return Number.isFinite(num) ? num : null;
      }

      function fmtDateLabel(raw) {
        if (!raw) return '';
        const str = String(raw).trim();
        const dt = new Date(str);
        if (!Number.isFinite(dt.getTime())) return str;
        const day = dt.getDate();
        const mon = dt.toLocaleDateString(undefined, { month: 'short' });
        const yr = dt.toLocaleDateString(undefined, { year: '2-digit' });
        return `${day} ${mon} ${yr}`;
      }

      function asDate(raw) {
        if (!raw) return null;
        const dt = new Date(raw);
        return Number.isFinite(dt.getTime()) ? dt : null;
      }

      function cloneCsvRows(sourceRows) {
        if (!Array.isArray(sourceRows)) return [];
        return sourceRows.map((row) => (Array.isArray(row) ? row.slice() : []));
      }

      function prepareCsvEditor(sourceRows) {
        const rows = cloneCsvRows(sourceRows);
        if (!rows.length) {
          return { headers: [], rows: [], indexes: {}, map: {} };
        }

        const headers = rows[0];
        for (let i = 0; i < headers.length; i += 1) {
          headers[i] = String(headers[i] ?? '');
        }

        const map = {};
        headers.forEach((header, index) => {
          const key = normHeader(header);
          if (key) {
            map[key] = index;
          }
        });

        const indexes = {};
        CSV_EDITOR_COLUMN_SPECS.forEach((spec) => {
          let targetIndex = null;
          for (let i = 0; i < spec.match.length; i += 1) {
            const key = spec.match[i];
            if (map[key] != null) {
              targetIndex = map[key];
              break;
            }
          }
          if (targetIndex != null) {
            indexes[spec.field] = targetIndex;
            return;
          }

          const normalized = normHeader(spec.header);
          const colIndex = headers.length;
          headers.push(spec.header);
          map[normalized] = colIndex;
          rows.forEach((row, idx) => {
            if (!Array.isArray(row)) return;
            if (row.length <= colIndex) {
              row.length = colIndex + 1;
            }
            if (idx === 0) {
              row[colIndex] = spec.header;
            } else if (row[colIndex] == null) {
              row[colIndex] = '';
            }
          });
          indexes[spec.field] = colIndex;
        });

        return {
          headers: headers.slice(),
          rows,
          indexes,
          map: { ...map }
        };
      }

      function csvEditorCellValue(row, index) {
        if (!Array.isArray(row) || index == null || index < 0) return '';
        const value = row[index];
        return value == null ? '' : String(value);
      }

      function assessCsvEditorRow(row, indexes) {
        const result = { fields: {}, hasError: false, hasWarning: false };
        const mark = (field, status) => {
          if (!field) return;
          const existing = result.fields[field];
          if (existing === 'error' || (existing === 'warning' && status === 'warning')) return;
          result.fields[field] = status;
          if (status === 'error') result.hasError = true;
          if (status === 'warning') result.hasWarning = true;
        };

        if (indexes.wedate != null) {
          const value = csvEditorCellValue(row, indexes.wedate).trim();
          if (!value || !asDate(value)) {
            mark('wedate', 'error');
          }
        }

        if (indexes.value1 != null) {
          const raw = csvEditorCellValue(row, indexes.value1);
          if (raw.trim() === '' || parseNumber(raw) == null) {
            mark('value1', 'error');
          }
        }

        CSV_EDITOR_WARNING_FIELDS.forEach((field) => {
          if (indexes[field] == null) return;
          const value = csvEditorCellValue(row, indexes[field]).trim();
          if (!value) {
            mark(field, 'warning');
          }
        });

        return result;
      }

      function buildCsvEditorRowHtml(row, rowIndex, indexes, assessment) {
        const rowClass = assessment.hasError ? 'table-danger' : (assessment.hasWarning ? 'table-warning' : '');
        const cells = CSV_EDITOR_FIELD_ORDER.map((field) => {
          const idx = indexes[field];
          const value = csvEditorCellValue(row, idx);
          const config = CSV_EDITOR_FIELD_CONFIG[field] || { type: 'text', placeholder: '' };
          const status = assessment.fields[field];
          let controlClass = 'csv-editor-input';
          if (status === 'error') {
            controlClass += ' is-invalid';
          } else if (status === 'warning') {
            controlClass += ' csv-editor-warning';
          }

          if (config.type === 'textarea') {
            return `<td><textarea class="form-control form-control-sm ${controlClass}" data-row-index="${rowIndex}" data-col-index="${idx}" data-field="${field}" rows="${config.rows || 2}" placeholder="${escapeHtml(config.placeholder)}">${escapeHtml(value)}</textarea></td>`;
          }

          const extras = [];
          if (config.step) extras.push(`step="${config.step}"`);
          if (config.inputmode) extras.push(`inputmode="${config.inputmode}"`);

          return `<td><input type="${config.type}" class="form-control form-control-sm ${controlClass}" data-row-index="${rowIndex}" data-col-index="${idx}" data-field="${field}" value="${escapeHtml(value)}" placeholder="${escapeHtml(config.placeholder)}" ${extras.join(' ')} /></td>`;
        }).join('');

        return `
              <tr class="${rowClass}" data-row-index="${rowIndex}">
                <td class="fw-semibold text-muted">${rowIndex}</td>
                ${cells}
              </tr>`;
      }

      function renderCsvEditorTable(focusState = null) {
        if (!csvEditorWorking || !model || !model.csvEditor || !csvEditorTableBody) return;
        const indexes = model.csvEditor.indexes || {};
        const showGaps = !!csvEditorFilters.showGaps;
        const searchTerm = (csvEditorFilters.search || '').trim().toLowerCase();
        const totalColumns = CSV_EDITOR_FIELD_ORDER.length + 1;
        const rows = csvEditorWorking;
        let html = '';
        let errorRows = 0;
        let warningOnlyRows = 0;
        let rendered = 0;

        for (let i = 1; i < rows.length; i += 1) {
          const row = rows[i] || [];
          const assessment = assessCsvEditorRow(row, indexes);
          if (assessment.hasError) errorRows += 1;
          else if (assessment.hasWarning) warningOnlyRows += 1;

          const needsAttention = assessment.hasError || assessment.hasWarning;
          if (showGaps && !needsAttention) continue;

          if (searchTerm) {
            const haystack = row
              .map((cell) => (cell == null ? '' : String(cell).toLowerCase()))
              .join(' ');
            if (!haystack.includes(searchTerm)) continue;
          }

          rendered += 1;
          html += buildCsvEditorRowHtml(row, i, indexes, assessment);
        }

        if (!rendered) {
          html = `<tr><td colspan="${totalColumns}" class="py-4 text-center text-muted">No rows match the current filters.</td></tr>`;
        }

        csvEditorTableBody.innerHTML = html;

        if (csvEditorRecordCount) csvEditorRecordCount.textContent = Math.max(rows.length - 1, 0);
        if (csvEditorErrorCount) csvEditorErrorCount.textContent = errorRows;
        if (csvEditorWarningCount) csvEditorWarningCount.textContent = warningOnlyRows;

        updateCsvEditorApplyState();

        if (focusState && focusState.rowIndex != null && focusState.field) {
          const selector = `[data-row-index="${focusState.rowIndex}"][data-field="${focusState.field}"]`;
          const target = csvEditorTableBody.querySelector(selector);
          if (target) {
            target.focus({ preventScroll: false });
            if (focusState.selectionStart != null && focusState.selectionEnd != null) {
              try {
                target.setSelectionRange(focusState.selectionStart, focusState.selectionEnd);
              } catch (err) {
                // ignore selection errors on unsupported inputs
              }
            }
          }
        }
      }

      function updateCsvEditorApplyState() {
        if (csvEditorApplyBtn) {
          csvEditorApplyBtn.disabled = !csvEditorDirty;
        }
      }

      function clearCsvEditorAlert() {
        if (csvEditorAlert) {
          csvEditorAlert.innerHTML = '';
        }
      }

      function setCsvEditorAlert(message, type = 'info') {
        if (!csvEditorAlert) return;
        if (!message) {
          csvEditorAlert.innerHTML = '';
          return;
        }
        csvEditorAlert.innerHTML = `
          <div class="alert alert-${type} alert-dismissible fade show" role="alert">
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
          </div>`;
      }

      function openCsvEditor() {
        if (!model || !model.csvEditor || !model.csvEditor.rows.length) {
          showStatus('Load a CSV file before opening the editor.', 'info');
          return;
        }
        csvEditorWorking = cloneCsvRows(model.csvEditor.rows);
        csvEditorDirty = false;
        csvEditorFilters.showGaps = !!model.graphOnly;
        csvEditorFilters.search = '';
        if (csvEditorShowGaps) csvEditorShowGaps.checked = csvEditorFilters.showGaps;
        if (csvEditorSearchInput) csvEditorSearchInput.value = '';
        clearCsvEditorAlert();
        renderCsvEditorTable();
        updateCsvEditorApplyState();
        if (!csvEditorModal && csvEditorModalEl && window.bootstrap) {
          csvEditorModal = new window.bootstrap.Modal(csvEditorModalEl);
        }
        if (csvEditorModal) {
          csvEditorModal.show();
        }
      }

      function handleCsvEditorInput(event) {
        const target = event.target;
        if (!target || !target.classList || !target.classList.contains('csv-editor-input')) return;
        if (!csvEditorWorking) return;
        const rowIndex = Number(target.dataset.rowIndex);
        const colIndex = Number(target.dataset.colIndex);
        if (Number.isNaN(rowIndex) || Number.isNaN(colIndex)) return;
        const value = target.value;
        if (!Array.isArray(csvEditorWorking[rowIndex])) {
          csvEditorWorking[rowIndex] = [];
        }
        csvEditorWorking[rowIndex][colIndex] = value;
        csvEditorDirty = true;
        const focusState = {
          rowIndex,
          field: target.dataset.field || '',
          selectionStart: target.selectionStart != null ? target.selectionStart : null,
          selectionEnd: target.selectionEnd != null ? target.selectionEnd : null
        };
        renderCsvEditorTable(focusState);
      }

      function validateCsvEditorWorking() {
        if (!csvEditorWorking || !model || !model.csvEditor) {
          return { errors: [], warnings: [] };
        }
        const indexes = model.csvEditor.indexes || {};
        const errors = [];
        const warnings = [];
        for (let i = 1; i < csvEditorWorking.length; i += 1) {
          const row = csvEditorWorking[i] || [];
          const assessment = assessCsvEditorRow(row, indexes);
          if (assessment.hasError) {
            const fields = Object.entries(assessment.fields)
              .filter(([, status]) => status === 'error')
              .map(([field]) => CSV_EDITOR_FIELD_LABELS[field] || field);
            errors.push({ rowIndex: i, fields });
          } else if (assessment.hasWarning) {
            const fields = Object.entries(assessment.fields)
              .filter(([, status]) => status === 'warning')
              .map(([field]) => CSV_EDITOR_FIELD_LABELS[field] || field);
            warnings.push({ rowIndex: i, fields });
          }
        }
        return { errors, warnings };
      }

      function applyCsvEditorChanges() {
        if (!model || !model.csvEditor || !csvEditorWorking) return;
        const validation = validateCsvEditorWorking();
        if (validation.errors.length) {
          const preview = validation.errors
            .slice(0, 5)
            .map((entry) => `Row ${entry.rowIndex}: ${escapeHtml(entry.fields.join(', '))}`)
            .join('<br>');
          setCsvEditorAlert(`Please correct the highlighted date or value fields before applying.<br>${preview}`, 'danger');
          csvEditorFilters.showGaps = true;
          if (csvEditorShowGaps) csvEditorShowGaps.checked = true;
          renderCsvEditorTable();
          return;
        }

        const updatedRows = cloneCsvRows(csvEditorWorking);
        const previousSettings = model && model.settings ? JSON.parse(JSON.stringify(model.settings)) : null;

        try {
          const rebuilt = buildModel(updatedRows);
          if (previousSettings) {
            const fallbackName = rebuilt.graphOnly
              ? (previousSettings.graphFallbackDivisionName ?? rebuilt.settings.graphFallbackDivisionName)
              : (rebuilt.settings.graphFallbackDivisionName ?? previousSettings.graphFallbackDivisionName ?? '');
            rebuilt.settings = {
              ...rebuilt.settings,
              defaultThemeKey: previousSettings.defaultThemeKey ?? rebuilt.settings.defaultThemeKey,
              defaultCustomAccent: previousSettings.defaultCustomAccent ?? rebuilt.settings.defaultCustomAccent,
              graphFallbackDivisionName: fallbackName
            };
          }

          model = rebuilt;
          model.csvEditor = prepareCsvEditor(updatedRows);
          csvEditorWorking = null;
          csvEditorDirty = false;
          clearCsvEditorAlert();
          updateCsvEditorApplyState();

          if (fallbackDivisionInput) {
            fallbackDivisionInput.value = model.settings?.graphFallbackDivisionName ?? '';
          }
          fallbackDivisionName = model.settings?.graphFallbackDivisionName ?? '';

          refreshDefaultThemeLabel();
          renderDivisions();
          generateBtn.disabled = !model || model.graphOnly;
          if (generateGraphsBtn) generateGraphsBtn.disabled = false;

          if (csvEditorModal) {
            csvEditorModal.hide();
          }

          showStatus('CSV edits applied. Regenerate the booklet or graphs to review the updates.', 'success');
        } catch (err) {
          setCsvEditorAlert(escapeHtml(err.message || 'Unable to rebuild data with the current edits.'), 'danger');
        }
      }

      function buildModel(rows) {
        const headers = rows[0].map((h) => String(h ?? ''));
        const map = {};
        const unnamedIndices = [];
        headers.forEach((header, index) => {
          const key = normHeader(header);
          if (key) {
            map[key] = index;
          } else {
            unnamedIndices.push(index);
          }
        });

        const idxStatbook = map['statbook'] ?? map['statbok'] ?? null;
        const idxStatDesc = map['statdescription'] ?? map['statdesc'] ?? map['statdescriptionlong'] ?? map['statdescriptionoptionalforprintingonly'] ?? null;
        const idxStatbookLong = map['statbooklong'] ?? map['statbooklongoptionalforprintingonly'] ?? map['statbooklongoptional'] ?? null;
        const idxShort = map['shortname'] ?? map['short'] ?? map['statbookshortname'];
        const idxDate = map['wedate'] ?? map['weekending'] ?? map['date'] ?? map['week'] ?? map['day'];
        const idxValue1 = map['value1'] ?? map['valueline1'] ?? map['value'];
        const idxValue2 = map['value2'] ?? map['valueline2'] ?? null;
        const idxValue3 = map['value3'] ?? map['valueline3'] ?? null;
        const idxValue4 = map['value4'] ?? map['valueline4'] ?? null;
        const idxStaff = map['staffmember'] ?? map['staff'] ?? null;
        const idxPost = map['post'] ?? map['postname'] ?? null;
        const idxDivision = map['division'] ?? map['div'] ?? null;
  const idxUpside = map['upsidedown'] ?? map['upsidedownflag'] ?? null;
  const idxMin = map['min'] ?? map['minimum'] ?? map['statmin'] ?? (unnamedIndices.length ? unnamedIndices[0] : null);
  const idxMax = map['max'] ?? map['maximum'] ?? map['statmax'] ?? (unnamedIndices.length > 1 ? unnamedIndices[1] : null);
        const idxPercent = map['ispercent'] ?? map['percent'] ?? null;
        const idxMoney = map['ismoney'] ?? map['money'] ?? null;
        const idxColor = map['colorflash'] ?? map['color'] ?? map['palette'] ?? null;

        if (idxDate == null || idxValue1 == null) {
          throw new Error('CSV must include WE-Date and Value1 columns.');
        }

        const hasDivisionColumn = idxDivision != null;
        const configuredFallback = (fallbackDivisionName ?? '').trim();
        const effectiveFallbackDivision = configuredFallback || (fallbackDivisionInput ? '' : 'General');

        const divisionMap = new Map();
        const divisionOrder = [];
        const globalDates = [];
        const carryForward = {
          division: '',
          staffName: '',
          postName: '',
          statbook: '',
          statdesc: '',
          statbookLong: '',
          shortName: '',
          upsideDown: null,
          min: null,
          max: null
        };

        const readTextValue = (row, idx, key) => {
          if (idx == null || idx >= row.length) {
            return carryForward[key] ?? '';
          }
          const raw = row[idx];
          const value = raw == null ? '' : String(raw).trim();
          if (!value) {
            return carryForward[key] ?? '';
          }
          carryForward[key] = value;
          return value;
        };

        const readNumericValue = (row, idx, key) => {
          if (idx == null || idx >= row.length) {
            return carryForward[key] ?? null;
          }
          const raw = row[idx];
          if (raw == null || raw === '') {
            return carryForward[key] ?? null;
          }
          const num = parseNumber(raw);
          if (num == null) {
            return carryForward[key] ?? null;
          }
          carryForward[key] = num;
          return num;
        };

        const readBoolValue = (row, idx, key, defaultValue = false) => {
          if (idx == null || idx >= row.length) {
            return carryForward[key] != null ? carryForward[key] : defaultValue;
          }
          const raw = row[idx];
          const str = raw == null ? '' : String(raw).trim();
          if (!str) {
            return carryForward[key] != null ? carryForward[key] : defaultValue;
          }
          const value = toBool(str);
          carryForward[key] = value;
          return value;
        };

        let lastStatFingerprint = null;

        for (let r = 1; r < rows.length; r += 1) {
          const row = rows[r];
          if (!row || !row.length) continue;
          const hasContent = row.some((cell) => cell != null && String(cell).trim() !== '');
          if (!hasContent) continue;

          const statbook = idxStatbook != null ? readTextValue(row, idxStatbook, 'statbook') : readTextValue(row, null, 'statbook');
          const statDesc = idxStatDesc != null ? readTextValue(row, idxStatDesc, 'statdesc') : readTextValue(row, null, 'statdesc');
          const statbookLong = idxStatbookLong != null ? readTextValue(row, idxStatbookLong, 'statbookLong') : readTextValue(row, null, 'statbookLong');
          const shortName = idxShort != null ? readTextValue(row, idxShort, 'shortName') : readTextValue(row, null, 'shortName');

          const statFingerprint = [statbook, shortName, statDesc, statbookLong].filter(Boolean).join('||') || `row-${r}`;
          if (statFingerprint !== lastStatFingerprint) {
            carryForward.min = null;
            carryForward.max = null;
            carryForward.upsideDown = null;
            lastStatFingerprint = statFingerprint;
          }

          let divisionName = hasDivisionColumn ? readTextValue(row, idxDivision, 'division') : effectiveFallbackDivision;
          if (hasDivisionColumn && !divisionName) {
            divisionName = 'Unassigned';
          }
          if (!hasDivisionColumn) {
            carryForward.division = divisionName;
          }

          if (!divisionMap.has(divisionName)) {
            divisionMap.set(divisionName, {
              name: divisionName,
              synthetic: !hasDivisionColumn,
              graphOnly: !hasDivisionColumn,
              staffOrder: [],
              staffMap: new Map(),
              colorHints: [],
              staff: [],
              color: null,
              vfpText: !hasDivisionColumn ? '' : (divisionName ? VFP_TEXT[divisionName.toUpperCase().replace(/[^A-Z]/g, '')] || '' : ''),
              vfpScale: 'normal',
              vfpAlign: 'center',
              vfpWidth: 'wide',
              vfpFont: 'large',
              customAccent: null
            });
            divisionOrder.push(divisionName);
          }

          const division = divisionMap.get(divisionName);

          if (idxColor != null && row[idxColor]) {
            division.colorHints.push(row[idxColor]);
          }

          let staffName = hasDivisionColumn ? readTextValue(row, idxStaff, 'staffName') : '';
          let postName = hasDivisionColumn ? readTextValue(row, idxPost, 'postName') : '';

          let staffKey;
          if (hasDivisionColumn) {
            staffKey = `${staffName}||${postName}`;
          } else {
            const graphStaffName = effectiveFallbackDivision ? `${effectiveFallbackDivision} Charts` : 'Graph Charts';
            staffKey = '__graph__staff__';
            staffName = graphStaffName.trim() || 'Graph Charts';
            postName = '';
          }

          if (!division.staffMap.has(staffKey)) {
            const staffObj = {
              staffName,
              post: postName,
              statsMap: new Map(),
              statOrder: [],
              stats: [],
              selected: new Set()
            };
            division.staffMap.set(staffKey, staffObj);
            division.staffOrder.push(staffKey);
          }

          const staffObj = division.staffMap.get(staffKey);
          if (!hasDivisionColumn) {
            staffObj.staffName = staffName;
            staffObj.post = postName;
          }

          const displayName = statbookLong || statDesc || (statbook && shortName ? `${statbook} — ${shortName}` : shortName || statbook || 'Stat');
          const checkboxLabel = shortName || statbook || statbookLong || displayName;
          const statKeyBase = [statbook, shortName, statDesc || statbookLong].filter(Boolean).join(' • ');
          const statKey = statKeyBase || `STAT-${staffObj.statOrder.length + 1}`;

          if (!staffObj.statsMap.has(statKey)) {
            const statObj = {
              key: statKey,
              statbook,
              statbookLong,
              shortName,
              statdesc: statDesc,
              displayName,
              checkboxLabel,
              dates: [],
              dateObjs: [],
              values: [[], [], [], []],
              upsideDown: false,
              min: null,
              max: null,
              isPercent: false,
              isMoney: false
            };
            staffObj.statsMap.set(statKey, statObj);
            staffObj.statOrder.push(statKey);
          }

          const statObj = staffObj.statsMap.get(statKey);
          const rawDate = row[idxDate];
          statObj.dates.push(fmtDateLabel(rawDate));
          const dateObj = asDate(rawDate);
          statObj.dateObjs.push(dateObj);
          if (dateObj) globalDates.push(dateObj);

          const v1 = parseNumber(row[idxValue1]);
          const v2 = idxValue2 != null ? parseNumber(row[idxValue2]) : null;
          const v3 = idxValue3 != null ? parseNumber(row[idxValue3]) : null;
          const v4 = idxValue4 != null ? parseNumber(row[idxValue4]) : null;

          statObj.values[0].push(v1);
          statObj.values[1].push(v2);
          statObj.values[2].push(v3);
          statObj.values[3].push(v4);

          if (idxUpside != null) {
            statObj.upsideDown = readBoolValue(row, idxUpside, 'upsideDown', false);
          }
          if (idxMin != null) {
            const minValue = readNumericValue(row, idxMin, 'min');
            if (minValue != null) statObj.min = minValue;
          }
          if (idxMax != null) {
            const maxValue = readNumericValue(row, idxMax, 'max');
            if (maxValue != null) statObj.max = maxValue;
          }
          if (idxPercent != null) {
            statObj.isPercent = readBoolValue(row, idxPercent, 'isPercent', false);
          }
          if (idxMoney != null) {
            statObj.isMoney = readBoolValue(row, idxMoney, 'isMoney', false);
          }
        }

        const divisions = divisionOrder.map((name, idx) => {
          const division = divisionMap.get(name);
          const colorToken = division.colorHints[0] || name || `division-${idx + 1}`;
          const theme = chooseColor(colorToken, idx);
          division.colorKey = theme.key;
          division.color = { ...theme };
          if (division.customAccent) {
            const palette = paletteFromAccent(division.customAccent);
            if (palette) {
              division.color = { ...division.color, ...palette, key: division.colorKey };
              division.customAccent = palette.accent;
            } else {
              division.customAccent = null;
            }
          } else if (defaultCustomAccent) {
            const palette = paletteFromAccent(defaultCustomAccent);
            if (palette) {
              division.color = { ...division.color, ...palette, key: division.colorKey };
              division.customAccent = palette.accent;
            }
          }
          division.staff = division.staffOrder.map((key) => {
            const staff = division.staffMap.get(key);
            staff.stats = staff.statOrder.map((statKey) => staff.statsMap.get(statKey));
            if (!staff.selected || !(staff.selected instanceof Set)) {
              staff.selected = new Set();
            }
            staff.selected = new Set(staff.statOrder);
            return staff;
          });
          return division;
        });

        let earliest = null;
        let latest = null;
        globalDates.forEach((dt) => {
          if (!dt) return;
          if (!earliest || dt < earliest) earliest = dt;
          if (!latest || dt > latest) latest = dt;
        });

        return {
          divisions,
          earliest,
          latest,
          graphOnly: !hasDivisionColumn,
          settings: {
            defaultThemeKey,
            defaultCustomAccent,
            graphFallbackDivisionName: !hasDivisionColumn ? effectiveFallbackDivision : (fallbackDivisionName ?? '').trim()
          }
        };
      }

      function renderDivisions(preserveOpen = false) {
        if (!model || !model.divisions.length) {
          divisionContainer.innerHTML = `
            <div class="text-center text-muted py-5">
              <i class="bi bi-diagram-3 fs-1 mb-3"></i>
              <div class="fw-semibold">No stats selected yet.</div>
              <div class="small">Load a CSV to manage divisions, staff and stats.</div>
            </div>`;
          updateSummary();
          return;
        }

        const openIds = preserveOpen
          ? new Set(Array.from(divisionContainer.querySelectorAll('.accordion-collapse.show')).map((el) => el.id))
          : new Set();

        const graphOnlyNotice = model.graphOnly
          ? `<div class="alert alert-warning d-flex align-items-center gap-2" role="alert">
              <i class="bi bi-bar-chart-line"></i>
              <div>This CSV does not include a Division column, so Staff Meeting pages are disabled. Adjust the graph label above if needed.</div>
            </div>`
          : '';

        const html = model.divisions
          .map((division, di) => {
            const selectedStats = division.staff.reduce((acc, staff) => acc + staff.selected.size, 0);
            const staffCount = division.staff.length;
            const badgeText = `${selectedStats} stat${selectedStats !== 1 ? 's' : ''}`;
            const swatch = `<button type="button" class="color-chip-btn me-2" data-action="division-color" data-di="${di}" title="Change division color" aria-label="Change ${escapeHtml(division.name)} division color" style="background:${division.color.base};border-color:${division.color.border}"></button>`;

            const staffHtml = division.staff
              .map((staff, si) => {
                const collapseId = `collapse-${di}-${si}`;
                const headingId = `heading-${di}-${si}`;
                const isOpen = openIds.has(collapseId) || (si === 0 && division.staff.length <= 3);
                const staffTitle = staff.staffName || staff.post || 'Unnamed Staff';
                const staffSubtitle = staff.post && staff.staffName ? `${staff.post} • ${staff.staffName}` : staff.post || staff.staffName || '';
                const statCountId = `statCount-${di}-${si}`;

                const statsHtml = staff.stats
                  .map((stat, ti) => {
                    const primaryTitle = resolveStatTitle(stat);
                    const secondaryTitle = resolveStatSubtitle(stat);
                    return `
                          <div class="form-check" draggable="true" data-drag-type="stat" data-di="${di}" data-si="${si}" data-stat-index="${ti}" data-drag-handle="stat">
                            <span class="drag-handle" data-drag-handle="stat" title="Drag to reorder stats"><i class="bi bi-grip-vertical"></i></span>
                            <input class="form-check-input stat-checkbox" type="checkbox" value="${escapeHtml(stat.key)}" data-di="${di}" data-si="${si}" data-index="${ti}" data-key="${escapeHtml(stat.key)}" ${
                      staff.selected.has(stat.key) ? 'checked' : ''
                    } />
                        <label class="form-check-label stat-label">
                          <span>${escapeHtml(primaryTitle)}</span>
                          ${secondaryTitle ? `<small>${escapeHtml(secondaryTitle)}</small>` : ''}
                        </label>
                      </div>`
                  })
                  .join('');

                return `
                  <div class="accordion-item" data-di="${di}" data-si="${si}" data-drag-type="staff" data-index="${si}" data-drag-handle="staff" draggable="true">
                    <h2 class="accordion-header" id="${headingId}">
                      <button class="accordion-button ${isOpen ? '' : 'collapsed'}" type="button" data-bs-toggle="collapse" data-bs-target="#${collapseId}" aria-expanded="${isOpen}" aria-controls="${collapseId}">
                        <div class="d-flex align-items-start gap-2 w-100">
                          <span class="drag-handle mt-1" data-drag-handle="staff" title="Drag to reorder staff"><i class="bi bi-grip-vertical"></i></span>
                          <div class="d-flex flex-column flex-grow-1">
                            <span class="fw-semibold">${escapeHtml(staffTitle)}</span>
                            ${staffSubtitle ? `<small class="text-muted">${escapeHtml(staffSubtitle)}</small>` : ''}
                          </div>
                        </div>
                      </button>
                    </h2>
                    <div id="${collapseId}" class="accordion-collapse collapse ${isOpen ? 'show' : ''}" aria-labelledby="${headingId}" data-bs-parent="#accordion-${di}">
                      <div class="accordion-body">
                        <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-3">
                          <span class="badge badge-light" id="${statCountId}">${staff.selected.size}/${staff.stats.length} selected</span>
                          <div class="btn-group btn-group-sm">
                            <button type="button" class="btn btn-outline-secondary" data-action="staff-up" data-di="${di}" data-si="${si}">
                              <i class="bi bi-arrow-up"></i>
                            </button>
                            <button type="button" class="btn btn-outline-secondary" data-action="staff-down" data-di="${di}" data-si="${si}">
                              <i class="bi bi-arrow-down"></i>
                            </button>
                          </div>
                        </div>
                        <div class="d-flex justify-content-end gap-2 mb-2 flex-wrap">
                          <button type="button" class="btn btn-outline-success btn-sm" data-action="select-all" data-di="${di}" data-si="${si}">Select all</button>
                          <button type="button" class="btn btn-outline-danger btn-sm" data-action="clear-stats" data-di="${di}" data-si="${si}">Clear</button>
                        </div>
                        <div class="stat-grid">
                          ${statsHtml}
                        </div>
                      </div>
                    </div>
                  </div>`;
              })
              .join('');

            return `
              <div class="card division-card" data-di="${di}" data-drag-type="division" data-index="${di}" data-drag-handle="division" draggable="true" style="border-left-color:${division.color.accent}">
                <div class="card-header bg-white d-flex justify-content-between align-items-center">
                  <div class="d-flex align-items-center gap-2">
                    <span class="drag-handle" data-drag-handle="division" title="Drag to reorder divisions"><i class="bi bi-grip-vertical"></i></span>
                    ${swatch}
                    <div>
                      <div class="fw-semibold text-uppercase">${escapeHtml(division.name)}</div>
                      <small class="text-muted">${staffCount} staff • ${badgeText}</small>
                    </div>
                  </div>
                  <div class="d-flex align-items-center gap-2">
                    <button type="button" class="btn btn-outline-secondary btn-sm" data-action="division-settings" data-di="${di}" title="Customize division">
                      <i class="bi bi-sliders"></i>
                    </button>
                    <div class="btn-group division-controls">
                      <button type="button" class="btn btn-outline-secondary btn-sm" data-action="div-up" data-di="${di}">
                        <i class="bi bi-arrow-up"></i>
                      </button>
                      <button type="button" class="btn btn-outline-secondary btn-sm" data-action="div-down" data-di="${di}">
                        <i class="bi bi-arrow-down"></i>
                      </button>
                    </div>
                  </div>
                </div>
                <div class="card-body">
                  ${division.staff.length ? `<div class="accordion" id="accordion-${di}">${staffHtml}</div>` : '<div class="text-muted">No staff records in this division.</div>'}
                </div>
              </div>`;
          })
          .join('');

        divisionContainer.innerHTML = graphOnlyNotice + html;
        updateSummary();
      }

      function handleCsvChange(event) {
        const file = event.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const text = e.target.result;
            clearStatus();
            const parsed = Papa.parse(text, { skipEmptyLines: true });
            if (!parsed.data || parsed.data.length < 2) {
              throw new Error('CSV must have a header row and at least one data row.');
            }
            model = buildModel(parsed.data);
            model.csvEditor = prepareCsvEditor(parsed.data);
            if (!model.settings) {
              model.settings = { defaultThemeKey, defaultCustomAccent };
            }
            if (model.settings.defaultThemeKey && COLOR_KEYS.includes(model.settings.defaultThemeKey)) {
              defaultThemeKey = model.settings.defaultThemeKey;
              if (defaultThemeSelect) {
                defaultThemeSelect.value = defaultThemeKey;
              }
            }
            if (model.settings.defaultCustomAccent) {
              const palette = paletteFromAccent(model.settings.defaultCustomAccent);
              defaultCustomAccent = palette ? palette.accent : null;
            } else {
              defaultCustomAccent = null;
            }
            model.settings.defaultThemeKey = defaultThemeKey;
            model.settings.defaultCustomAccent = defaultCustomAccent;
            if (model.settings.graphFallbackDivisionName != null) {
              const appliedFallback = String(model.settings.graphFallbackDivisionName || '').trim();
              if (fallbackDivisionInput) {
                fallbackDivisionInput.value = appliedFallback;
              }
              fallbackDivisionName = appliedFallback;
              if (model.graphOnly) {
                const graphStaffLabel = (appliedFallback ? `${appliedFallback} Charts` : 'Graph Charts').trim() || 'Graph Charts';
                model.divisions.forEach((division) => {
                  division.name = appliedFallback;
                  division.staff.forEach((staff) => {
                    staff.staffName = graphStaffLabel;
                    staff.post = '';
                  });
                });
              }
            } else if (model.graphOnly && fallbackDivisionInput) {
              const appliedFallback = (fallbackDivisionInput.value ?? '').trim();
              fallbackDivisionName = appliedFallback;
              model.settings.graphFallbackDivisionName = appliedFallback;
              const graphStaffLabel = (appliedFallback ? `${appliedFallback} Charts` : 'Graph Charts').trim() || 'Graph Charts';
              model.divisions.forEach((division) => {
                division.name = appliedFallback;
                division.staff.forEach((staff) => {
                  staff.staffName = graphStaffLabel;
                  staff.post = '';
                });
              });
            }
            refreshDefaultThemeLabel();
            renderDivisions();
            generateBtn.disabled = !model || model.graphOnly;
            if (generateGraphsBtn) generateGraphsBtn.disabled = false;
            if (editCsvBtn) editCsvBtn.disabled = false;
            csvEditorWorking = null;
            csvEditorDirty = false;
            csvEditorFilters.showGaps = !!model.graphOnly;
            csvEditorFilters.search = '';
            if (csvEditorShowGaps) csvEditorShowGaps.checked = csvEditorFilters.showGaps;
            if (csvEditorSearchInput) csvEditorSearchInput.value = '';
            clearCsvEditorAlert();
            updateCsvEditorApplyState();
            const modeNote = model.graphOnly ? ' <span class="badge bg-warning text-dark ms-2">Graph mode only</span>' : '';
            showStatus(`Loaded CSV file: <strong>${escapeHtml(file.name)}</strong>${modeNote}`, 'success');
          } catch (err) {
            model = null;
            renderDivisions();
            generateBtn.disabled = true;
            if (generateGraphsBtn) generateGraphsBtn.disabled = true;
            if (editCsvBtn) editCsvBtn.disabled = true;
            csvEditorWorking = null;
            csvEditorDirty = false;
            clearCsvEditorAlert();
            updateCsvEditorApplyState();
            showStatus(err.message || 'Unable to parse CSV file.', 'danger');
          }
        };
        reader.onerror = () => {
          showStatus('Unable to read the selected CSV file.', 'danger');
        };
        reader.readAsText(file);
      }

      function handleLogoChange(event) {
        const file = event.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
          logoDataURL = e.target.result;
          logoPreview.src = logoDataURL;
          logoPreview.classList.remove('d-none');
          logoPlaceholder.classList.add('d-none');
        };
        reader.onerror = () => showStatus('Unable to read the selected logo file.', 'warning');
        reader.readAsDataURL(file);
      }

      function handleDivisionClick(event) {
        const button = event.target.closest('button[data-action]');
        if (!button || !model) return;
        const action = button.dataset.action;
        const di = Number(button.dataset.di);
        const si = Number(button.dataset.si);

        if (action === 'division-color') {
          openDivisionSettings(di, 'color');
        } else if (action === 'div-up' && di > 0) {
          [model.divisions[di - 1], model.divisions[di]] = [model.divisions[di], model.divisions[di - 1]];
          renderDivisions(true);
        } else if (action === 'div-down' && di < model.divisions.length - 1) {
          [model.divisions[di + 1], model.divisions[di]] = [model.divisions[di], model.divisions[di + 1]];
          renderDivisions(true);
        } else if (action === 'division-settings') {
          openDivisionSettings(di);
        } else if ((action === 'staff-up' || action === 'staff-down') && !Number.isNaN(si)) {
          const division = model.divisions[di];
          if (!division) return;
          if (action === 'staff-up' && si > 0) {
            [division.staff[si - 1], division.staff[si]] = [division.staff[si], division.staff[si - 1]];
            renderDivisions(true);
          } else if (action === 'staff-down' && si < division.staff.length - 1) {
            [division.staff[si + 1], division.staff[si]] = [division.staff[si], division.staff[si + 1]];
            renderDivisions(true);
          }
        } else if (action === 'select-all' && !Number.isNaN(si)) {
          const division = model.divisions[di];
          if (!division) return;
          const staff = division.staff[si];
          staff.selected = new Set(staff.stats.map((stat) => stat.key));
          renderDivisions(true);
        } else if (action === 'clear-stats' && !Number.isNaN(si)) {
          const division = model.divisions[di];
          if (!division) return;
          const staff = division.staff[si];
          staff.selected.clear();
          renderDivisions(true);
        }
      }

      function handleDivisionChange(event) {
        const target = event.target;
        if (!target.classList.contains('stat-checkbox') || !model) return;
        const di = Number(target.dataset.di);
        const si = Number(target.dataset.si);
        const key = target.dataset.key;
        const division = model.divisions[di];
        if (!division) return;
        const staff = division.staff[si];
        if (!staff) return;
        if (target.checked) {
          staff.selected.add(key);
        } else {
          staff.selected.delete(key);
        }
        const badge = document.getElementById(`statCount-${di}-${si}`);
        if (badge) {
          badge.textContent = `${staff.selected.size}/${staff.stats.length} selected`;
        }
        updateSummary();
      }

      function openDivisionSettings(di, focusField = null) {
        if (!divisionModal || !model) return;
        const division = model.divisions[di];
        if (!division) return;
        editingDivisionIndex = di;
        divisionModalFocusField = focusField;
        if (divisionNameDisplay) {
          divisionNameDisplay.value = division.name;
        }
        if (divisionVfpInput) {
          divisionVfpInput.value = division.vfpText || '';
        }
        if (divisionColorSelect) {
          const key = division.colorKey || (division.color ? division.color.key : null) || COLOR_KEYS[0];
          divisionColorSelect.value = key;
        }
        if (divisionCustomColorInput) {
          if (divisionCustomColorInput.dataset.reset) {
            delete divisionCustomColorInput.dataset.reset;
          }
          const accent = normalizeHexColor(division.customAccent || (division.color ? division.color.accent : null)) || '#1D4ED8';
          divisionCustomColorInput.value = accent;
          divisionCustomColorInput.dataset.initial = accent;
        }
        if (divisionCustomColorDefault) {
          divisionCustomColorDefault.checked = false;
        }
        if (divisionCustomColorReset) {
          divisionCustomColorReset.disabled = !division.customAccent;
        }
        if (divisionVfpScaleSelect) {
          divisionVfpScaleSelect.value = division.vfpScale || 'normal';
        }
        if (divisionVfpAlignSelect) {
          divisionVfpAlignSelect.value = division.vfpAlign || 'center';
        }
        if (divisionVfpWidthSelect) {
          divisionVfpWidthSelect.value = division.vfpWidth || 'wide';
        }
        if (divisionVfpFontSelect) {
          divisionVfpFontSelect.value = division.vfpFont || 'large';
        }
        divisionModal.show();
      }

      let dragState = null;
      let currentDragTarget = null;

      function moveArrayItem(list, from, to) {
        if (from === to || from < 0 || to < 0 || from >= list.length || to >= list.length) return;
        const [item] = list.splice(from, 1);
        list.splice(to, 0, item);
      }

      function clearDragStyles() {
        if (currentDragTarget) {
          currentDragTarget.classList.remove('drag-target');
          currentDragTarget = null;
        }
        if (dragState && dragState.element) {
          dragState.element.classList.remove('dragging');
        }
      }

      function handleDragStart(event) {
        const handle = event.target.closest('[data-drag-handle]');
        if (!handle) {
          event.preventDefault();
          return;
        }
        const source = handle.closest('[data-drag-type]');
        if (!source) return;
        const type = source.dataset.dragType;
        dragState = {
          type,
          element: source,
          di: source.dataset.di != null ? Number(source.dataset.di) : null,
          si: source.dataset.si != null ? Number(source.dataset.si) : null,
          statIndex: source.dataset.statIndex != null ? Number(source.dataset.statIndex) : null,
          index: source.dataset.index != null ? Number(source.dataset.index) : null
        };
        source.classList.add('dragging');
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = 'move';
          event.dataTransfer.setData('text/plain', type);
        }
      }

      function handleDragOver(event) {
        if (!dragState) return;
        const target = event.target.closest('[data-drag-type]');
        if (!target || target.dataset.dragType !== dragState.type) return;
        event.preventDefault();
        if (currentDragTarget !== target) {
          if (currentDragTarget) {
            currentDragTarget.classList.remove('drag-target');
          }
          currentDragTarget = target;
          currentDragTarget.classList.add('drag-target');
        }
        if (event.dataTransfer) {
          event.dataTransfer.dropEffect = 'move';
        }
      }

      function handleDrop(event) {
        if (!dragState) return;
        const target = event.target.closest('[data-drag-type]');
        if (!target || target.dataset.dragType !== dragState.type) return;
        event.preventDefault();

        try {
          if (dragState.type === 'division') {
            const from = dragState.di ?? dragState.index ?? 0;
            const to = target.dataset.di != null ? Number(target.dataset.di) : Number(target.dataset.index);
            moveArrayItem(model.divisions, from, to);
            renderDivisions(true);
          } else if (dragState.type === 'staff') {
            const di = dragState.di;
            const targetDivision = Number(target.dataset.di);
            if (di == null || Number.isNaN(di) || di !== targetDivision) return;
            const division = model.divisions[di];
            if (!division) return;
            const from = dragState.si ?? dragState.index ?? 0;
            const to = target.dataset.si != null ? Number(target.dataset.si) : Number(target.dataset.index);
            moveArrayItem(division.staff, from, to);
            renderDivisions(true);
          } else if (dragState.type === 'stat') {
            const di = dragState.di;
            const si = dragState.si;
            if (di == null || si == null) return;
            const targetDi = Number(target.dataset.di);
            const targetSi = Number(target.dataset.si);
            if (di !== targetDi || si !== targetSi) return;
            const division = model.divisions[di];
            if (!division) return;
            const staff = division.staff[si];
            if (!staff) return;
            const from = dragState.statIndex ?? 0;
            const to = target.dataset.statIndex != null ? Number(target.dataset.statIndex) : from;
            moveArrayItem(staff.stats, from, to);
            renderDivisions(true);
          }
        } finally {
          clearDragStyles();
          dragState = null;
        }
      }

      function handleDragEnd() {
        clearDragStyles();
        dragState = null;
      }

      function buildSeriesPayload(stat) {
        const names = ['Primary', 'Secondary', 'Tertiary', 'Quaternary'];
        return stat.values
          .map((arr, idx) => ({
            name: names[idx],
            data: arr.map((v) => {
              if (v == null || v === '') return null;
              const numeric = typeof v === 'number' ? v : Number(v);
              return Number.isFinite(numeric) ? numeric : null;
            })
          }))
          .map((series) => {
            const finiteValues = (series.data || []).filter((v) => Number.isFinite(v));
            const min = finiteValues.length ? Math.min(...finiteValues) : null;
            const max = finiteValues.length ? Math.max(...finiteValues) : null;
            const decimals = finiteValues.reduce((acc, value) => Math.max(acc, countDecimals(value)), 0);
            return {
              ...series,
              meta: {
                hasData: finiteValues.length > 0,
                min,
                max,
                decimals
              }
            };
          })
          .filter((series) => series.meta && series.meta.hasData);
      }

      function preparePayloadBase() {
        if (!model || !model.divisions.length) {
          throw new Error('Load a CSV and select at least one stat before generating.');
        }

        const divisions = [];
        let earliest = null;
        let latest = null;

        model.divisions.forEach((division) => {
          const staffEntries = [];
          division.staff.forEach((staff) => {
            const chosenStats = staff.stats.filter((stat) => staff.selected.has(stat.key));
            if (!chosenStats.length) return;
            const statsPayload = chosenStats.map((stat) => {
              stat.dateObjs.forEach((dt) => {
                if (dt && (!earliest || dt < earliest)) earliest = dt;
                if (dt && (!latest || dt > latest)) latest = dt;
              });
              const shortTitle = getShortTitle(stat);
              const longTitle = getLongTitle(stat);
              const resolvedTitle = resolveStatTitle(stat);
              const unusedTitle = statTitleMode === 'short' ? longTitle : shortTitle;
              return {
                key: stat.key,
                title: resolvedTitle,
                displayName: resolvedTitle,
                statbook: stat.statbook,
                shortName: stat.shortName,
                statdesc: stat.statdesc,
                shortTitle,
                longTitle,
                alternateTitle: unusedTitle && unusedTitle !== resolvedTitle ? unusedTitle : '',
                dates: stat.dates,
                dateObjs: stat.dateObjs.map((d) => (d ? d.toISOString() : null)),
                series: buildSeriesPayload(stat),
                isPercent: stat.isPercent,
                isMoney: stat.isMoney,
                min: stat.min,
                max: stat.max,
                upsideDown: stat.upsideDown
              };
            });
            staffEntries.push({
              staffName: staff.staffName,
              post: staff.post,
              stats: statsPayload
            });
          });
          if (staffEntries.length) {
            divisions.push({
              name: division.name,
              synthetic: !!division.synthetic,
              graphOnly: !!division.graphOnly,
              color: division.color,
              vfpScale: division.vfpScale || 'normal',
              vfpAlign: division.vfpAlign || 'center',
              vfpWidth: division.vfpWidth || 'wide',
              vfpFont: division.vfpFont || 'large',
              customAccent: division.customAccent || null,
              vfpText: division.vfpText,
              staff: staffEntries
            });
          }
        });

        if (!divisions.length) {
          throw new Error('Select at least one stat to include.');
        }

        return { divisions, earliest, latest };
      }

      function buildBookletPayload() {
        const { divisions, earliest, latest } = preparePayloadBase();
        const title = titleInput.value.trim() || 'Weekly BP';
        const subtitle = subtitleInput.value.trim();
        return {
          title,
          subtitle,
          logo: logoDataURL,
          pageSize,
          earliest: earliest ? earliest.toISOString() : null,
          latest: latest ? latest.toISOString() : null,
          divisions
        };
      }

      function buildGraphPayload() {
        const { divisions, earliest, latest } = preparePayloadBase();
        const title = titleInput.value.trim() || 'Weekly BP';
        const subtitle = subtitleInput.value.trim();
        const payload = {
          title,
          subtitle,
          headingMode: graphHeadingMode,
          layoutMode: graphLayoutMode,
          logoMode: graphLogoMode,
          pageSize,
          earliest: earliest ? earliest.toISOString() : null,
          latest: latest ? latest.toISOString() : null,
          divisions
        };
        if (graphLogoMode === 'show' && logoDataURL) {
          payload.logo = logoDataURL;
        }
        return payload;
      }

      function buildSettingsSnapshot() {
        if (!model || !model.divisions.length) {
          throw new Error('Load a CSV and configure at least one division before saving settings.');
        }
        return {
          version: 1,
          savedAt: new Date().toISOString(),
          title: titleInput.value.trim() || null,
          subtitle: subtitleInput.value.trim() || null,
          pageSize,
          statTitleMode,
          graphHeadingMode,
          graphLogoMode,
          graphLayoutMode,
          defaultThemeKey,
          defaultCustomAccent,
          graphOnly: !!(model && model.graphOnly),
          graphFallbackDivisionName: model?.settings?.graphFallbackDivisionName ?? (fallbackDivisionName ?? ''),
          logo: logoDataURL || null,
          divisions: model.divisions.map((division) => ({
            name: division.name,
            colorKey: division.colorKey,
            customAccent: division.customAccent || null,
            vfpText: division.vfpText,
            vfpScale: division.vfpScale,
            vfpAlign: division.vfpAlign,
            vfpWidth: division.vfpWidth,
            vfpFont: division.vfpFont || 'medium',
            graphOnly: !!division.graphOnly,
            synthetic: !!division.synthetic
          }))
        };
      }

      function downloadSettings() {
        try {
          const snapshot = buildSettingsSnapshot();
          const json = JSON.stringify(snapshot, null, 2);
          const blob = new Blob([json], { type: 'application/json' });
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          const stamp = new Date().toISOString().split('T')[0];
          link.href = url;
          link.download = `booklet-settings-${stamp}.json`;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
          showStatus('Settings downloaded as JSON.', 'success');
        } catch (error) {
          showStatus(error.message || 'Unable to prepare settings for download.', 'danger');
        }
      }

      const sharedChildRuntimeSource = () => {
        if (window.__BOOKLET_SHARED_RUNTIME__) return;
        window.__BOOKLET_SHARED_RUNTIME__ = (() => {
          const TREND_UP = '#0f172a';
          const TREND_DOWN = '#b91c1c';
          const CM_TO_PX = 37.7952755906;
          const MARKER_TICK_LENGTH_PX = Math.round(1 * CM_TO_PX);
          const LABEL_CLEARANCE_PX = Math.round(0.3 * CM_TO_PX);
          const LABEL_RISER_OFFSET_PX = MARKER_TICK_LENGTH_PX + LABEL_CLEARANCE_PX;
          const DROPLINE_COLOR = 'rgba(15,23,42,0.22)';
          const DROPLINE_TICK_COLOR = 'rgba(15,23,42,0.32)';
          const measurementConstants = {
            CM_TO_PX,
            MARKER_TICK_LENGTH_PX,
            LABEL_CLEARANCE_PX,
            LABEL_RISER_OFFSET_PX
          };

          const applyDataLabelStyles = (chart) => {
            if (!chart || !chart.series) return;
            const plotTop = Number.isFinite(chart.plotTop) ? chart.plotTop : 0;
            const plotLeft = Number.isFinite(chart.plotLeft) ? chart.plotLeft : 0;
            const HighchartsInstance = window.Highcharts;
            const renderer = chart && chart.renderer && typeof chart.renderer.g === 'function'
              ? chart.renderer
              : null;
            let tickGroup = null;
            if (renderer) {
              tickGroup = chart._markerTickLayer;
              if (!tickGroup) {
                tickGroup = renderer.g('marker-ticks').attr({ zIndex: 5 }).add();
                chart._markerTickLayer = tickGroup;
              }
              if (Array.isArray(chart._markerTickGraphics)) {
                chart._markerTickGraphics.forEach((graphic) => {
                  if (graphic && typeof graphic.destroy === 'function') {
                    graphic.destroy();
                  }
                });
              }
              chart._markerTickGraphics = [];
            }
            (chart.series || []).forEach((series) => {
              if (!series.visible) return;
              const baseColor = '#0f172a';
              (series.points || []).forEach((point) => {
                const dataLabel = point && point.dataLabel;
                if (!dataLabel) return;
                dataLabel.css({
                  color: baseColor,
                  textOutline: 'none',
                  fontWeight: 600,
                  backgroundColor: 'transparent',
                  padding: '0'
                });
                if (dataLabel.element) {
                  dataLabel.element.setAttribute('text-anchor', 'middle');
                  dataLabel.element.style.background = 'none';
                  dataLabel.element.style.border = 'none';
                  dataLabel.element.style.boxShadow = 'none';
                  dataLabel.element.style.filter = 'none';
                  dataLabel.element.style.textShadow = '0 0 2px rgba(255,255,255,0.82), 0 0 5px rgba(255,255,255,0.6)';
                }
                if (!Number.isFinite(point.plotX) || !Number.isFinite(point.plotY)) return;
                const axis = series && series.yAxis;
                const axisMax = axis && Number.isFinite(axis.max) ? axis.max
                  : (axis && Number.isFinite(axis.dataMax) ? axis.dataMax : null);
                const pointValue = Number.isFinite(point.y) ? point.y : null;
                const atAxisMax = axisMax != null && pointValue != null && pointValue >= axisMax - 1e-9;
                if (atAxisMax) {
                  if (typeof dataLabel.hide === 'function') {
                    dataLabel.hide();
                  } else if (dataLabel.element) {
                    dataLabel.element.style.visibility = 'hidden';
                  }
                  return;
                }
                if (typeof dataLabel.show === 'function') {
                  dataLabel.show();
                } else if (dataLabel.element) {
                  dataLabel.element.style.visibility = '';
                }
                const anchorX = plotLeft + point.plotX;
                const markerY = plotTop + point.plotY;
                const anchorY = markerY - (MARKER_TICK_LENGTH_PX + LABEL_CLEARANCE_PX);
                const bbox = typeof dataLabel.getBBox === 'function' ? dataLabel.getBBox() : null;
                const rawHeight = bbox && Number.isFinite(bbox.height) ? bbox.height : (dataLabel.height || 0);
                const MIN_LABEL_SPAN_PX = 26;
                const verticalSpan = Math.max(MIN_LABEL_SPAN_PX, rawHeight || 0);
                const clearance = Math.max(7, LABEL_CLEARANCE_PX * 0.32);
                const currentX = Number.isFinite(dataLabel.x)
                  ? dataLabel.x
                  : (Number.isFinite(dataLabel.translateX)
                    ? dataLabel.translateX
                    : (Number.isFinite(anchorX) ? anchorX : 0));
                const offsetY = Number.isFinite(anchorY) ? anchorY : 0;

                //const finalY = offsetY - LABEL_RISER_OFFSET_PX - clearance;

                const finalY = offsetY - verticalSpan - clearance;
                if (typeof dataLabel.translate === 'function') {
                   dataLabel.translate(currentX, finalY);
                } else {
                  dataLabel.attr({ x: currentX, y: finalY });
                }
                dataLabel.x = currentX;
                dataLabel.y = finalY;
                dataLabel.alignAttr = dataLabel.alignAttr || {};
                dataLabel.alignAttr.x = currentX;
                dataLabel.alignAttr.y = finalY;
                dataLabel.placed = true;
                if (typeof dataLabel.updateTransform === 'function') {
                  dataLabel.updateTransform();
                }

                if (tickGroup && Number.isFinite(anchorX) && Number.isFinite(markerY)) {
                  const tickTop = Math.max(plotTop, markerY - MARKER_TICK_LENGTH_PX);
                  const tickPath = ['M', anchorX, tickTop, 'L', anchorX, markerY];
                  const tickGraphic = renderer
                    .path(tickPath)
                    .attr({
                      stroke: DROPLINE_TICK_COLOR,
                      'stroke-width': 1.6,
                      'stroke-linecap': 'round'
                    })
                    .add(tickGroup);
                  chart._markerTickGraphics.push(tickGraphic);
                }

                // --- Uniform label alignment for all numeric values ---
            // const fixedOffset = LABEL_RISER_OFFSET_PX + 8; // base offset above tick (adjust ±8 to taste)
            // const finalY = offsetY - fixedOffset;

            // // Force a consistent baseline (no per-character wobble)
            // if (dataLabel.element) {
            //   dataLabel.element.style.dominantBaseline = 'central';
            //   dataLabel.element.style.alignmentBaseline = 'middle';
            // }

            // if (typeof dataLabel.translate === 'function') {
            //   dataLabel.translate(currentX, finalY);
            // } else {
            //   dataLabel.attr({ x: currentX, y: finalY });
            // }
            // dataLabel.x = currentX;
            // dataLabel.y = finalY;
            // dataLabel.alignAttr = { x: currentX, y: finalY };
            // dataLabel.placed = true;
            // if (typeof dataLabel.updateTransform === 'function') {
            //   dataLabel.updateTransform();
            // }

              });
            });
          };

          const installDroplines = (chart) => {
            if (!chart || !chart.renderer) return;
            const storeKey = '_droplineGraphics';
            chart[storeKey] = chart[storeKey] || [];
            chart[storeKey].forEach((graphic) => {
              if (graphic && typeof graphic.destroy === 'function') {
                graphic.destroy();
              }
            });
            chart[storeKey] = [];
            const plotTop = chart.plotTop;
            const plotLeft = chart.plotLeft;
            const plotHeight = chart.plotHeight;
            const bottomY = plotTop + plotHeight;
            (chart.series || []).forEach((series) => {
              if (!series.visible) return;
              (series.points || []).forEach((point) => {
                if (point.isNull || !Number.isFinite(point.plotX) || !Number.isFinite(point.plotY)) return;
                const x = plotLeft + point.plotX;
                const markerY = plotTop + point.plotY;
                const dropPath = ['M', x, markerY, 'L', x, bottomY];
                const dropLine = chart.renderer
                  .path(dropPath)
                  .attr({ stroke: DROPLINE_COLOR, 'stroke-width': 1.2, zIndex: 1, 'stroke-dasharray': '2.5 4' })
                  .add();
                chart[storeKey].push(dropLine);
                // Marker-top tick marks disabled to avoid visual clutter around labels.
              });
            });
          };

          const ensureChartDecorators = (() => {
            let attached = false;
            return () => {
              const HighchartsInstance = window.Highcharts;
              if (!HighchartsInstance || typeof HighchartsInstance.addEvent !== 'function' || attached) return;
              attached = true;
              HighchartsInstance.addEvent(HighchartsInstance.Chart, 'render', function onRender() {
                applyDataLabelStyles(this);
                installDroplines(this);
              });
            };
          })();

          const clamp = (value, min, max) => {
            if (!Number.isFinite(value)) return value;
            return Math.max(min, Math.min(max, value));
          };

          const escapeHtml = (str) => String(str || '').replace(/[&<>"'`]/g, (c) => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;',
            '`': '&#96;'
          }[c] || c));

          const isPlainObject = (value) => value && typeof value === 'object' && !Array.isArray(value);

          const deepMerge = (target, source) => {
            const base = Array.isArray(target) ? target.slice() : { ...target };
            if (!source) return base;
            if (Array.isArray(source)) {
              return source.slice();
            }
            if (!isPlainObject(source)) {
              return source;
            }
            Object.keys(source).forEach((key) => {
              const value = source[key];
              if (Array.isArray(value)) {
                base[key] = value.slice();
              } else if (isPlainObject(value)) {
                base[key] = deepMerge(base[key] || {}, value);
              } else {
                base[key] = value;
              }
            });
            return base;
          };

          const countDecimals = (value) => {
            if (!Number.isFinite(value)) return 0;
            const str = String(value).toLowerCase();
            if (str.includes('e')) {
              const [base, exponent] = str.split('e');
              const decimalsPart = (base.split('.')[1] || '').replace(/0+$/, '');
              const exp = Number(exponent);
              return Math.max(0, decimalsPart.length - exp);
            }
            if (!str.includes('.')) return 0;
            return str.split('.')[1].replace(/0+$/, '').length;
          };

          const roundToDecimals = (value, decimals) => {
            const safeDecimals = Math.max(0, Math.min(decimals, 8));
            const factor = Math.pow(10, safeDecimals + 2);
            return Math.round(value * factor) / factor;
          };

          const niceStep = (rawStep, constraints = {}) => {
            const allowDecimalTicks = !!constraints.allowDecimalTicks;
            const minimum = allowDecimalTicks
              ? Math.max(Number.EPSILON, constraints.minimum || Number.EPSILON)
              : Math.max(1, constraints.minimum || 1);
            if (!Number.isFinite(rawStep) || rawStep === 0) return minimum;
            const absStep = Math.abs(rawStep);
            const exponent = Math.floor(Math.log10(absStep));
            const magnitude = Math.pow(10, exponent);
            const fraction = absStep / magnitude;
            let niceFraction;
            if (fraction <= 1) niceFraction = 1;
            else if (fraction <= 2.4) niceFraction = 2;
            else if (fraction <= 3.2) niceFraction = 2.5;
            else if (fraction <= 5) niceFraction = 5;
            else niceFraction = 10;
            let step = niceFraction * magnitude;
            if (constraints.isPercent) {
              const percentChoices = allowDecimalTicks
                ? [0.001, 0.002, 0.003, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.25, 0.5, 1, 2, 3, 5, 10, 20, 25, 50]
                : [1, 2, 5, 10, 20, 25, 50];
              const adjusted = percentChoices.find((choice) => choice >= step) || percentChoices[percentChoices.length - 1];
              step = Math.max(adjusted, minimum);
            } else {
              step = Math.max(step, minimum);
            }
            if (!allowDecimalTicks && !Number.isInteger(step)) {
              step = Math.ceil(step);
            }
            return rawStep < 0 ? -step : step;
          };

          const formatNumberForDisplay = (value, decimals, HighchartsInstance, options = {}) => {
            if (!Number.isFinite(value)) return '';
            const { isPercent = false, isMoney = false, disableAbbrev = false } = options;
            const valueDecimals = countDecimals(value);
            const requestedDecimals = Number.isFinite(decimals) ? Math.max(0, decimals) : 0;
            const targetPrecision = Math.max(requestedDecimals, valueDecimals);
            const thresholds = [
              { limit: 1_000_000_000, divisor: 1_000_000_000, suffix: ' B' },
              { limit: 1_000_000, divisor: 1_000_000, suffix: ' M' },
              { limit: 1_000, divisor: 1_000, suffix: ' K' }
            ];
            const abs = Math.abs(value);
            let scaled = value;
            let suffix = '';
            let dp = Math.max(0, Math.min(4, targetPrecision));

            const smallMagnitude = abs < 1;
            if (!smallMagnitude) {
              if (isPercent) {
                const percentCap = Math.max(4, requestedDecimals);
                dp = Math.min(dp, percentCap);
              } else if (isMoney) {
                dp = Math.min(dp, Math.max(2, requestedDecimals));
              } else {
                dp = Math.min(dp, Math.max(2, requestedDecimals));
              }
            } else if (isPercent) {
              const percentCap = Math.max(4, requestedDecimals);
              dp = Math.min(dp, percentCap);
            }

            if (!disableAbbrev && !isPercent && abs >= 1_000) {
              const threshold = thresholds.find((entry) => abs >= entry.limit);
              if (threshold) {
                scaled = value / threshold.divisor;
                suffix = threshold.suffix;
                if (!smallMagnitude) {
                  dp = Math.min(dp, Math.max(2, requestedDecimals));
                }
              }
            }

            let formatted;
            if (!HighchartsInstance || typeof HighchartsInstance.numberFormat !== 'function') {
              formatted = dp === 0
                ? Math.round(scaled).toString()
                : Number(scaled.toFixed(dp)).toString();
            } else {
              formatted = HighchartsInstance.numberFormat(scaled, dp);
            }
            if (dp > 0) {
              formatted = formatted.replace(/(\.\d*?[1-9])0+$/g, '$1');
              formatted = formatted.replace(/\.0+$/g, '');
            }
            return `${formatted}${suffix}`.trim();
          };

          const deriveAxisStats = (stat) => {
            if (!stat || !Array.isArray(stat.series) || !stat.series.length) {
              return [stat];
            }
            const limit = Math.min(stat.series.length, 4);
            const stats = [];
            for (let i = 0; i < limit; i += 1) {
              if (i === 0) {
                stats.push(stat);
                continue;
              }
              const seriesEntry = stat.series[i];
              if (!seriesEntry) continue;
              const meta = seriesEntry.meta || {};
              const axisStat = {
                ...stat,
                series: [seriesEntry],
                min: Number.isFinite(meta.min) ? meta.min : stat.min,
                max: Number.isFinite(meta.max) ? meta.max : stat.max
              };
              stats.push(axisStat);
            }
            return stats.length ? stats : [stat];
          };

          const MANUAL_SCALE_STORAGE_KEY = 'booklet.manualScaleOverrides';

          const supportsManualScaleStorage = (() => {
            try {
              if (typeof window === 'undefined' || !window.localStorage) return false;
              const testKey = '__booklet_scale_test__';
              window.localStorage.setItem(testKey, '1');
              window.localStorage.removeItem(testKey);
              return true;
            } catch (err) {
              return false;
            }
          })();

          const sanitizeManualOverrideEntry = (entry) => {
            if (!entry || typeof entry !== 'object') return {};
            const result = {};
            Object.keys(entry).forEach((axisKey) => {
              const axisEntry = entry[axisKey];
              const min = Number(axisEntry?.min);
              const max = Number(axisEntry?.max);
              if (Number.isFinite(min) && Number.isFinite(max) && max > min) {
                result[axisKey] = { min, max };
              }
            });
            return result;
          };

          const loadManualScaleStore = () => {
            if (!supportsManualScaleStorage) return {};
            try {
              const raw = window.localStorage.getItem(MANUAL_SCALE_STORAGE_KEY);
              if (!raw) return {};
              const parsed = JSON.parse(raw);
              if (!parsed || typeof parsed !== 'object') return {};
              const output = {};
              Object.keys(parsed).forEach((chartKey) => {
                const sanitized = sanitizeManualOverrideEntry(parsed[chartKey]);
                if (Object.keys(sanitized).length) {
                  output[chartKey] = sanitized;
                }
              });
              return output;
            } catch (err) {
              return {};
            }
          };

          const persistManualScaleStore = (registry) => {
            if (!supportsManualScaleStorage) return;
            try {
              const payload = {};
              registry.forEach((value, key) => {
                if (!key) return;
                const sanitized = sanitizeManualOverrideEntry(value);
                if (Object.keys(sanitized).length) {
                  payload[key] = sanitized;
                }
              });
              const keys = Object.keys(payload);
              if (!keys.length) {
                window.localStorage.removeItem(MANUAL_SCALE_STORAGE_KEY);
              } else {
                window.localStorage.setItem(MANUAL_SCALE_STORAGE_KEY, JSON.stringify(payload));
              }
            } catch (err) {
              /* ignore storage errors */
            }
          };

          const initialManualScaleState = loadManualScaleStore();
          const manualScaleRegistry = new Map();
          Object.keys(initialManualScaleState).forEach((chartKey) => {
            manualScaleRegistry.set(chartKey, initialManualScaleState[chartKey]);
          });

          const updateManualScaleRegistry = (chartKey, overrides) => {
            if (!chartKey) return {};
            const sanitized = sanitizeManualOverrideEntry(overrides);
            if (Object.keys(sanitized).length) {
              manualScaleRegistry.set(chartKey, sanitized);
            } else {
              manualScaleRegistry.delete(chartKey);
            }
            persistManualScaleStore(manualScaleRegistry);
            return sanitized;
          };

          const fetchManualScaleOverrides = (chartKey) => {
            if (!chartKey) return {};
            const stored = manualScaleRegistry.get(chartKey);
            const sanitized = sanitizeManualOverrideEntry(stored);
            if (Object.keys(sanitized).length) {
              if (stored !== sanitized) {
                manualScaleRegistry.set(chartKey, sanitized);
                persistManualScaleStore(manualScaleRegistry);
              }
              return sanitized;
            }
            if (stored) {
              manualScaleRegistry.delete(chartKey);
              persistManualScaleStore(manualScaleRegistry);
            }
            return {};
          };

          let scalingModalEl = null;
          let scalingModalInstance = null;
          let scalingModalAxesContainer = null;
          let scalingModalApplyBtn = null;
          let scalingModalResetBtn = null;
          let activeScaleContext = null;
          const SCALE_AXIS_LABELS = ['Primary', 'Secondary', 'Tertiary', 'Quaternary'];

          const alignAxisBoundsToPrimary = (boundsList) => {
            if (!Array.isArray(boundsList) || !boundsList.length) {
              return Array.isArray(boundsList) ? boundsList : [];
            }
            const primary = boundsList[0];
            if (!primary) {
              return boundsList.slice();
            }
            const cloneBounds = (bounds) => {
              if (!bounds) return null;
              const ticks = Array.isArray(bounds.ticks) ? bounds.ticks.slice() : [];
              return { ...bounds, ticks };
            };
            const primaryClone = cloneBounds(primary);
            return boundsList.map((bounds, index) => {
              if (index === 0) {
                return primaryClone;
              }
              const target = cloneBounds(bounds) || primaryClone;
              if (!primaryClone) {
                return target;
              }
              return {
                ...target,
                min: primaryClone.min,
                max: primaryClone.max,
                ticks: Array.isArray(primaryClone.ticks) ? primaryClone.ticks.slice() : [],
                allowDecimals: primaryClone.allowDecimals,
                decimals: primaryClone.decimals,
                reversed: primaryClone.reversed
              };
            });
          };

          const ensureSoftWhiteHaloFilter = () => {
            if (typeof document === 'undefined') return;
            if (document.getElementById('softWhiteHalo')) return;
            const svgNS = 'http://www.w3.org/2000/svg';
            const svg = document.createElementNS(svgNS, 'svg');
            svg.setAttribute('aria-hidden', 'true');
            svg.style.position = 'absolute';
            svg.style.width = '0';
            svg.style.height = '0';
            svg.style.overflow = 'hidden';
            svg.innerHTML = `
              <filter id="softWhiteHalo" x="-50%" y="-50%" width="50%" height="50%">
                <feGaussianBlur in="SourceAlpha" stdDeviation="2.0" result="blur"></feGaussianBlur>
                <feMerge>
                  <feMergeNode in="blur"></feMergeNode>
                  <feMergeNode in="SourceGraphic"></feMergeNode>
                </feMerge>
              </filter>`;
            document.body.appendChild(svg);
          };

          const ensureFallbackScalingStyles = () => {
            if (typeof document === 'undefined') return;
            if (document.getElementById('niceScalingModalStyles')) return;
            const style = document.createElement('style');
            style.id = 'niceScalingModalStyles';
            style.textContent = [
              '.nice-scaling-modal {',
              '  position: fixed;',
              '  inset: 0;',
              '  display: none;',
              '  align-items: center;',
              '  justify-content: center;',
              '  padding: 1.5rem;',
              '  background: rgba(15, 23, 42, 0.4);',
              '  backdrop-filter: blur(8px);',
              '  z-index: 1080;',
              '  font-family: "Inter", "Segoe UI", system-ui, sans-serif;',
              '}',
              '.nice-scaling-modal.show {',
              '  display: flex;',
              '}',
              '.nice-scaling-modal .modal-dialog {',
              '  margin: 0;',
              '  width: min(720px, 92vw);',
              '  pointer-events: auto;',
              '  display: flex;',
              '  justify-content: center;',
              '}',
              '.nice-scaling-modal .modal-content {',
              '  border-radius: 18px;',
              '  pointer-events: auto;',
              '  width: 100%;',
              '  max-height: min(85vh, 720px);',
              '  overflow: hidden;',
              '  border: 1px solid rgba(148, 163, 184, 0.35);',
              '  box-shadow: 0 40px 90px -32px rgba(15, 23, 42, 0.55);',
              '  background: linear-gradient(165deg, #ffffff, #f1f5f9);',
              '  color: #0f172a;',
              '  display: flex;',
              '  flex-direction: column;',
              '}',
              '.nice-scaling-modal .modal-header,',
              '.nice-scaling-modal .modal-footer {',
              '  background: rgba(248, 250, 252, 0.9);',
              '  border-color: rgba(148, 163, 184, 0.35);',
              '  padding: 1.25rem 1.75rem;',
              '}',
              '.nice-scaling-modal .modal-header .modal-title {',
              '  margin: 0;',
              '  font-size: 1.05rem;',
              '  letter-spacing: 0.04em;',
              '  text-transform: uppercase;',
              '  font-weight: 600;',
              '}',
              '.nice-scaling-modal .modal-close {',
              '  background: transparent;',
              '  border: none;',
              '  color: rgba(15, 23, 42, 0.55);',
              '  font-size: 1.2rem;',
              '  cursor: pointer;',
              '  line-height: 1;',
              '  padding: 0.25rem;',
              '  border-radius: 999px;',
              '  transition: color 0.2s ease, background 0.2s ease;',
              '}',
              '.nice-scaling-modal .modal-close:hover {',
              '  color: #1d4ed8;',
              '  background: rgba(59, 130, 246, 0.1);',
              '}',
              '.nice-scaling-modal .modal-body {',
              '  background: transparent;',
              '  padding: 1.5rem 1.75rem;',
              '  overflow-y: auto;',
              '}',
              '.nice-scaling-modal .modal-body .modal-subtitle {',
              '  margin: 0 0 1rem;',
              '  font-size: 0.82rem;',
              '  color: rgba(15, 23, 42, 0.65);',
              '  letter-spacing: 0.03em;',
              '}',
              '.nice-scaling-modal [data-role="axis-container"] {',
              '  display: grid;',
              '  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));',
              '  gap: 1.1rem;',
              '  padding: 0;',
              '  margin: 0;',
              '}',
              '.nice-scaling-axis {',
              '  list-style: none;',
              '  margin: 0;',
              '  padding: 1.1rem 1.15rem;',
              '  border-radius: 14px;',
              '  border: 1px solid rgba(203, 213, 225, 0.7);',
              '  background: #ffffff;',
              '  box-shadow: 0 18px 36px -28px rgba(15, 23, 42, 0.55);',
              '  display: flex;',
              '  flex-direction: column;',
              '  gap: 0.85rem;',
              '  color: #0f172a;',
              '}',
              '.nice-scaling-axis.has-custom {',
              '  border-color: rgba(59, 130, 246, 0.55);',
              '  box-shadow: 0 24px 52px -28px rgba(59, 130, 246, 0.45);',
              '}',
              '.nice-scaling-axis-header {',
              '  display: flex;',
              '  justify-content: space-between;',
              '  align-items: flex-start;',
              '  gap: 1rem;',
              '}',
              '.nice-scaling-axis-header .axis-label {',
              '  font-size: 0.78rem;',
              '  font-weight: 700;',
              '  letter-spacing: 0.08em;',
              '  text-transform: uppercase;',
              '  color: #0f172a;',
              '}',
              '.nice-scaling-axis-header .axis-range {',
              '  margin-top: 0.25rem;',
              '  font-size: 0.78rem;',
              '  color: rgba(15, 23, 42, 0.72);',
              '}',
              '.nice-scaling-axis .scale-chip {',
              '  display: inline-flex;',
              '  align-items: center;',
              '  gap: 0.35rem;',
              '  padding: 0.25rem 0.65rem;',
              '  border-radius: 999px;',
              '  background: rgba(96, 165, 250, 0.12);',
              '  color: #1d4ed8;',
              '  font-size: 0.68rem;',
              '  font-weight: 600;',
              '  letter-spacing: 0.08em;',
              '  text-transform: uppercase;',
              '}',
              '.nice-scaling-axis .scale-chip.is-custom {',
              '  background: rgba(59, 130, 246, 0.2);',
              '  color: #1e3a8a;',
              '}',
              '.nice-scaling-axis .row {',
              '  display: grid;',
              '  grid-template-columns: repeat(2, minmax(0, 1fr));',
              '  gap: 0.75rem;',
              '}',
              '.nice-scaling-axis .form-label {',
              '  font-size: 0.68rem;',
              '  text-transform: uppercase;',
              '  letter-spacing: 0.08em;',
              '  font-weight: 600;',
              '  color: rgba(15, 23, 42, 0.7);',
              '  margin-bottom: 0.35rem;',
              '}',
              '.nice-scaling-axis input[type="number"] {',
              '  width: 100%;',
              '  border-radius: 10px;',
              '  border: 1px solid rgba(148, 163, 184, 0.6);',
              '  padding: 0.45rem 0.65rem;',
              '  font-size: 0.85rem;',
              '  font-weight: 500;',
              '  color: #0f172a;',
              '  background: rgba(255, 255, 255, 0.95);',
              '  transition: border-color 0.2s ease, box-shadow 0.2s ease;',
              '}',
              '.nice-scaling-axis input[type="number"]:focus {',
              '  outline: none;',
              '  border-color: rgba(59, 130, 246, 0.75);',
              '  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);',
              '}',
              '.nice-scaling-axis .axis-meta {',
              '  font-size: 0.75rem;',
              '  color: rgba(15, 23, 42, 0.6);',
              '  display: flex;',
              '  flex-direction: column;',
              '  gap: 0.35rem;',
              '}',
              '.nice-scaling-axis .axis-meta strong {',
              '  color: #0f172a;',
              '  font-weight: 600;',
              '}',
              '.nice-scaling-modal .modal-footer {',
              '  display: flex;',
              '  flex-wrap: wrap;',
              '  gap: 0.65rem;',
              '  justify-content: flex-end;',
              '  align-items: center;',
              '}',
              '.nice-scaling-modal .modal-footer .btn {',
              '  border-radius: 999px;',
              '  padding: 0.45rem 1.1rem;',
              '  border: 1px solid rgba(148, 163, 184, 0.55);',
              '  background: #ffffff;',
              '  color: #0f172a;',
              '  font-size: 0.78rem;',
              '  font-weight: 600;',
              '  cursor: pointer;',
              '  transition: background 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease;',
              '}',
              '.nice-scaling-modal .modal-footer .btn-outline {',
              '  background: transparent;',
              '}',
              '.nice-scaling-modal .modal-footer .btn-primary {',
              '  background: linear-gradient(135deg, #2563eb, #1d4ed8);',
              '  border-color: transparent;',
              '  color: #f8fafc;',
              '}',
              '.nice-scaling-modal .modal-footer .btn-primary:hover {',
              '  box-shadow: 0 12px 32px -18px rgba(37, 99, 235, 0.65);',
              '}',
              '.nice-scaling-modal .modal-footer .btn:hover {',
              '  border-color: rgba(37, 99, 235, 0.6);',
              '  background: rgba(59, 130, 246, 0.08);',
              '}',
              '@media (max-width: 640px) {',
              '  .nice-scaling-modal {',
              '    padding: 1rem;',
              '  }',
              '  .nice-scaling-modal .modal-dialog {',
              '    width: min(96vw, 520px);',
              '  }',
              '  .nice-scaling-modal [data-role="axis-container"] {',
              '    grid-template-columns: 1fr;',
              '  }',
              '}'
            ].join('\n');
            document.head.appendChild(style);
          };

          const ensureScalingModal = () => {
            if (scalingModalEl) return;
            const wrapper = document.createElement('div');
            wrapper.innerHTML = `
              <div class="modal fade nice-scaling-modal" id="scaleModal" tabindex="-1" aria-hidden="true">
                <div class="modal-dialog modal-dialog-centered">
                  <div class="modal-content">
                    <div class="modal-header">
                      <h6 class="modal-title">Adjust Chart Scales</h6>
                      <button type="button" class="modal-close" data-bs-dismiss="modal" aria-label="Close">&#215;</button>
                    </div>
                    <div class="modal-body">
                      <p class="modal-subtitle">Fine-tune any axis. Leave fields empty to keep the automatic scaling.</p>
                      <div class="nice-scaling-axes" data-role="axis-container"></div>
                    </div>
                    <div class="modal-footer">
                      <button type="button" class="btn btn-outline" data-bs-dismiss="modal">Cancel</button>
                      <button type="button" class="btn btn-outline" data-role="reset" id="resetScale">Reset to Auto</button>
                      <button type="button" class="btn btn-primary" data-role="apply" id="applyScale">Apply &amp; Close</button>
                    </div>
                  </div>
                </div>
              </div>`;
            scalingModalEl = wrapper.firstElementChild;
            scalingModalAxesContainer = scalingModalEl.querySelector('[data-role="axis-container"]');
            scalingModalApplyBtn = scalingModalEl.querySelector('[data-role="apply"]');
            scalingModalResetBtn = scalingModalEl.querySelector('[data-role="reset"]');
            document.body.appendChild(scalingModalEl);

            const bootstrapModal = window.bootstrap && window.bootstrap.Modal ? window.bootstrap.Modal : null;
            if (bootstrapModal) {
              scalingModalInstance = new bootstrapModal(scalingModalEl);
              scalingModalEl.addEventListener('hidden.bs.modal', () => {
                activeScaleContext = null;
                if (scalingModalAxesContainer) {
                  scalingModalAxesContainer.innerHTML = '';
                }
              });
            } else {
              ensureFallbackScalingStyles();
              scalingModalEl.classList.add('modal', 'nice-scaling-modal');
              scalingModalInstance = {
                show() {
                  if (document.body && scalingModalEl.parentElement !== document.body) {
                    document.body.appendChild(scalingModalEl);
                  }
                  scalingModalEl.style.display = 'flex';
                  scalingModalEl.classList.add('show');
                },
                hide() {
                  scalingModalEl.style.display = 'none';
                  scalingModalEl.classList.remove('show');
                  if (scalingModalEl.parentElement && scalingModalEl.parentElement !== document.body && document.body) {
                    document.body.appendChild(scalingModalEl);
                  }
                  activeScaleContext = null;
                  if (scalingModalAxesContainer) {
                    scalingModalAxesContainer.innerHTML = '';
                  }
                }
              };
            }

            const dismissButtons = scalingModalEl.querySelectorAll('[data-bs-dismiss]');
            dismissButtons.forEach((btn) => {
              btn.addEventListener('click', () => {
                if (scalingModalInstance && !bootstrapModal) {
                  scalingModalInstance.hide();
                }
              });
            });

            if (scalingModalApplyBtn) {
              scalingModalApplyBtn.addEventListener('click', () => {
                if (!activeScaleContext) {
                  scalingModalInstance.hide();
                  return;
                }
                const overrides = {};
                const axisNodes = scalingModalAxesContainer ? scalingModalAxesContainer.querySelectorAll('.nice-scaling-axis') : [];
                axisNodes.forEach((axisNode) => {
                  const axisIndex = Number(axisNode.getAttribute('data-axis-index'));
                  const minInput = axisNode.querySelector('input[data-role="min"]');
                  const maxInput = axisNode.querySelector('input[data-role="max"]');
                  const minValue = minInput ? minInput.value.trim() : '';
                  const maxValue = maxInput ? maxInput.value.trim() : '';
                  if (minValue === '' && maxValue === '') return;
                  const parsedMin = Number(minValue);
                  const parsedMax = Number(maxValue);
                  if (!Number.isFinite(parsedMin) || !Number.isFinite(parsedMax) || parsedMax <= parsedMin) return;
                  overrides[axisIndex] = { min: parsedMin, max: parsedMax };
                });
                const sanitizedOverrides = updateManualScaleRegistry(activeScaleContext.chartKey, overrides);
                applyManualOverrides(activeScaleContext.chart, sanitizedOverrides, activeScaleContext.context);
                scalingModalInstance.hide();
              });
            }

            if (scalingModalResetBtn) {
              scalingModalResetBtn.addEventListener('click', () => {
                if (!activeScaleContext) {
                  scalingModalInstance.hide();
                  return;
                }
                updateManualScaleRegistry(activeScaleContext.chartKey, null);
                applyAutoScaling(activeScaleContext.chart, activeScaleContext.context);
                scalingModalInstance.hide();
              });
            }
          };

          const rebindChartSeries = (chart, context, overrides = null) => {
            if (!chart || !chart.series || !chart.series.length) return;
            const seriesList = chart.series;
            const axisCount = chart.yAxis ? chart.yAxis.length : 1;
            const contextOrigins = Array.isArray(context?.seriesOrigins) ? context.seriesOrigins : null;
            const storedOrigins = Array.isArray(chart.__seriesOrigins) ? chart.__seriesOrigins : null;
            const origins = seriesList.map((_, idx) => {
              const origin = contextOrigins && Number.isFinite(contextOrigins[idx]) ? contextOrigins[idx]
                : storedOrigins && Number.isFinite(storedOrigins[idx]) ? storedOrigins[idx]
                : idx;
              return Math.max(0, Math.min(origin, axisCount - 1));
            });

            const effectiveOverrides = overrides != null ? overrides : fetchManualScaleOverrides(context?.chartKey);

            const visibleAxisIndex = (() => {
              if (!chart.yAxis || !chart.yAxis.length) return 0;
              const found = chart.yAxis.findIndex((axis) => axis && axis.visible !== false);
              return found >= 0 ? found : 0;
            })();

            seriesList.forEach((series, idx) => {
              const originIdx = origins[idx] ?? 0;
              let targetAxisIdx = originIdx;
              if (!Number.isFinite(targetAxisIdx) || targetAxisIdx <= 0) {
                targetAxisIdx = visibleAxisIndex;
              }
              const axisRef = chart.yAxis && chart.yAxis[targetAxisIdx];
              const hasOverrideForAxis = effectiveOverrides && Object.prototype.hasOwnProperty.call(effectiveOverrides, targetAxisIdx);
              if (!axisRef || (axisRef.visible === false && !hasOverrideForAxis)) {
                targetAxisIdx = visibleAxisIndex;
              }
              const currentAxisIdx = series && series.yAxis ? series.yAxis.index : 0;
              if (currentAxisIdx !== targetAxisIdx) {
                series.update({ yAxis: targetAxisIdx }, false);
                series.redraw();
              }
            });
          };

          const applyAutoScaling = (chart, context) => {
            if (!chart || !context) return;
            const yAxes = chart.yAxis || [];
            if (!yAxes.length) return;
            const allowMultipleAxes = !!(context.axisOptions && context.axisOptions.allowMultipleAxes);
            const axisStats = allowMultipleAxes ? deriveAxisStats(context.stat).slice(0, 4) : [context.stat];
            const rawAutoBounds = axisStats.map((axisStat) => computeAxisBounds(axisStat, context.axisOptions || {}));
            const autoBoundsList = alignAxisBoundsToPrimary(rawAutoBounds);
            context.axisStats = axisStats;
            context.autoBounds = autoBoundsList;
            yAxes.forEach((axis, index) => {
              const bounds = autoBoundsList[index] || autoBoundsList[0];
              if (!bounds) return;
              const isPrimary = index === 0;
              const axisOptions = axis && axis.options ? axis.options : {};
              const baseLabels = axisOptions.labels || {};
              const resolvedLineWidth = isPrimary
                ? (Number.isFinite(axisOptions.lineWidth) ? Math.max(1, axisOptions.lineWidth) : 2)
                : 0;
              const resolvedGridWidth = isPrimary
                ? (Number.isFinite(axisOptions.gridLineWidth) ? Math.max(0, axisOptions.gridLineWidth) : 1)
                : 0;
              axis.update({
                min: bounds.min,
                max: bounds.max,
                tickPositions: bounds.ticks,
                startOnTick: true,
                endOnTick: true,
                allowDecimals: bounds.allowDecimals !== false,
                visible: isPrimary,
                opposite: !isPrimary && index % 2 === 1,
                labels: {
                  ...baseLabels,
                  align: baseLabels.align != null ? baseLabels.align : (isPrimary ? 'right' : 'left'),
                  enabled: isPrimary
                },
                gridLineWidth: resolvedGridWidth,
                lineWidth: resolvedLineWidth,
                reversed: !!bounds.reversed
              }, false);
            });
            if (yAxes.length > axisStats.length) {
              for (let i = axisStats.length; i < yAxes.length; i += 1) {
                const extraAxis = yAxes[i];
                if (!extraAxis) continue;
                extraAxis.update({
                  visible: false,
                  opposite: false,
                  labels: { enabled: false },
                  tickPositions: null,
                  min: null,
                  max: null,
                  allowDecimals: true,
                  gridLineWidth: 0,
                  lineWidth: 0,
                  reversed: false
                }, false);
              }
            }
            const primaryAxis = yAxes[0];
            if (primaryAxis) {
              const primaryOptions = primaryAxis.options || {};
              primaryAxis.update({
                visible: true,
                opposite: false,
                labels: {
                  ...(primaryOptions.labels || {}),
                  enabled: true,
                  align: (primaryOptions.labels && primaryOptions.labels.align) || 'right'
                },
                lineWidth: Number.isFinite(primaryOptions.lineWidth) ? Math.max(1, primaryOptions.lineWidth) : 2,
                gridLineWidth: Number.isFinite(primaryOptions.gridLineWidth) ? Math.max(0, primaryOptions.gridLineWidth) : 1
              }, false);
            }
            rebindChartSeries(chart, context);
            chart.redraw();
          };

          const applyManualOverrides = (chart, overrides, context) => {
            if (!chart || !context) return;
            const yAxes = chart.yAxis || [];
            if (!yAxes.length) return;
            const effectiveOverrides = overrides != null ? overrides : fetchManualScaleOverrides(context.chartKey);
            const allowMultipleAxes = !!(context.axisOptions && context.axisOptions.allowMultipleAxes);
            const axisStats = context.axisStats || (allowMultipleAxes ? deriveAxisStats(context.stat).slice(0, 4) : [context.stat]);
            const storedAutoBounds = Array.isArray(context.autoBounds) ? context.autoBounds : [];
            const autoBoundsList = storedAutoBounds.length ? alignAxisBoundsToPrimary(storedAutoBounds) : [];
            const desiredTickCountFallback = Math.max(5, Math.min((context.axisOptions && context.axisOptions.desiredTickCount) || 6, 10));
            yAxes.forEach((axis, index) => {
              const axisStat = axisStats[index] || axisStats[0] || context.stat;
              const axisOptions = axis && axis.options ? axis.options : {};
              const baseLabels = axisOptions.labels || {};
              const defaultOpposite = typeof axisOptions.opposite === 'boolean'
                ? axisOptions.opposite
                : (index > 0 && index % 2 === 1);
              const resolvedLineWidth = index === 0
                ? (Number.isFinite(axisOptions.lineWidth) ? Math.max(1, axisOptions.lineWidth) : 2)
                : 0;
              const resolvedGridWidth = index === 0
                ? (Number.isFinite(axisOptions.gridLineWidth) ? Math.max(0, axisOptions.gridLineWidth) : 1)
                : 0;
              const labelAlign = baseLabels.align != null ? baseLabels.align : (index === 0 ? 'right' : 'left');
              const labelOptions = {
                ...baseLabels,
                align: labelAlign,
                enabled: index === 0
              };
              const override = effectiveOverrides[index];
              if (override && Number.isFinite(override.min) && Number.isFinite(override.max) && override.max > override.min) {
                  const span = override.max - override.min;
                  const desiredTickBase = Number.isFinite(axisOptions.desiredTickCount)
                    ? Number(axisOptions.desiredTickCount)
                    : desiredTickCountFallback;
                  const desiredSegments = Math.max(1, Math.round(desiredTickBase) - 1);
                  const minInterval = Number.isFinite(axisOptions.minimumTickInterval) && axisOptions.minimumTickInterval > 0
                    ? axisOptions.minimumTickInterval
                    : Number.EPSILON;
                  const manualStep = niceStep(span / Math.max(desiredSegments, 1), {
                    allowDecimalTicks: true,
                    minimum: minInterval,
                    isPercent: !!axisStat?.isPercent
                  });
                  let manualSegments = Number.isFinite(manualStep) && manualStep > 0
                    ? Math.round(span / manualStep)
                    : desiredSegments;
                  manualSegments = Math.max(4, Math.min(16, manualSegments || desiredSegments || 4));
                  const manualBounds = computeAxisBounds(axisStat, {
                    ...(context.axisOptions || {}),
                    min: override.min,
                    max: override.max,
                    padFraction: 0,
                    manualDivisions: manualSegments
                  });
                  const ticks = manualBounds && Array.isArray(manualBounds.ticks) && manualBounds.ticks.length
                    ? manualBounds.ticks
                    : [override.min, override.max];
                  const minValue = manualBounds && Number.isFinite(manualBounds.min) ? manualBounds.min : override.min;
                  const maxValue = manualBounds && Number.isFinite(manualBounds.max) ? manualBounds.max : override.max;
                  axis.update({
                    min: minValue,
                    max: maxValue,
                    tickPositions: ticks,
                    startOnTick: true,
                    endOnTick: true,
                    allowDecimals: manualBounds ? manualBounds.allowDecimals !== false : undefined,
                    visible: index === 0,
                    opposite: defaultOpposite,
                    labels: labelOptions,
                    gridLineWidth: resolvedGridWidth,
                    lineWidth: resolvedLineWidth,
                    reversed: !!manualBounds.reversed
                  }, false);
                  return;
              }

              if (index === 0) {
                const auto = autoBoundsList[0]
                  ? autoBoundsList[0]
                  : computeAxisBounds(axisStat, context.axisOptions || {});
                axis.update({
                  min: auto.min,
                  max: auto.max,
                  tickPositions: auto.ticks,
                  startOnTick: true,
                  endOnTick: true,
                  allowDecimals: auto.allowDecimals !== false,
                  visible: true,
                  labels: {
                    ...labelOptions,
                    enabled: true
                  },
                  gridLineWidth: resolvedGridWidth,
                  lineWidth: resolvedLineWidth,
                  reversed: !!auto.reversed
                }, false);
                const refreshedAutoBounds = autoBoundsList.length ? autoBoundsList : [auto];
                context.autoBounds = alignAxisBoundsToPrimary(refreshedAutoBounds);
              } else {
                const fallbackSource = autoBoundsList[index]
                  || autoBoundsList[0]
                  || computeAxisBounds(axisStat, context.axisOptions || {});
                const fallback = alignAxisBoundsToPrimary([autoBoundsList[0] || fallbackSource])[0] || fallbackSource;
                axis.update({
                  min: fallback.min,
                  max: fallback.max,
                  tickPositions: fallback.ticks,
                  startOnTick: true,
                  endOnTick: true,
                  allowDecimals: fallback.allowDecimals !== false,
                  visible: index === 0,
                  opposite: defaultOpposite,
                  labels: labelOptions,
                  gridLineWidth: 0,
                  lineWidth: 0,
                  reversed: !!fallback.reversed
                }, false);
              }
            });
            const primaryAxis = yAxes[0];
            if (primaryAxis) {
              const primaryOptions = primaryAxis.options || {};
              primaryAxis.update({
                visible: true,
                opposite: false,
                labels: {
                  ...(primaryOptions.labels || {}),
                  enabled: true,
                  align: (primaryOptions.labels && primaryOptions.labels.align) || 'right'
                },
                lineWidth: Number.isFinite(primaryOptions.lineWidth) ? Math.max(1, primaryOptions.lineWidth) : 2,
                gridLineWidth: Number.isFinite(primaryOptions.gridLineWidth) ? Math.max(0, primaryOptions.gridLineWidth) : 1
              }, false);
            }
            rebindChartSeries(chart, context, effectiveOverrides);
            chart.redraw();
          };

          const openScalingModal = (chart, context) => {
            ensureScalingModal();
            if (!scalingModalEl || !scalingModalInstance || !scalingModalAxesContainer) return;
            activeScaleContext = {
              chart,
              context,
              chartKey: context.chartKey
            };
            const usingBootstrap = !!(window.bootstrap && window.bootstrap.Modal);
            if (usingBootstrap) {
              if (scalingModalEl.parentElement !== document.body) {
                document.body.appendChild(scalingModalEl);
                scalingModalEl.removeAttribute('data-embedded');
              }
            } else if (document.body && scalingModalEl.parentElement !== document.body) {
              document.body.appendChild(scalingModalEl);
              scalingModalEl.removeAttribute('data-embedded');
            }
            scalingModalAxesContainer.innerHTML = '';
            const overrides = fetchManualScaleOverrides(context.chartKey);
            const yAxes = chart.yAxis || [];
            const seriesList = (context && context.stat && Array.isArray(context.stat.series)) ? context.stat.series : [];
            yAxes.slice(0, 4).forEach((axis, index) => {
              const seriesName = seriesList[index] && seriesList[index].name ? seriesList[index].name : null;
              const axisLabel = seriesName || SCALE_AXIS_LABELS[index] || `Axis ${index + 1}`;
              const axisNode = document.createElement('div');
              axisNode.className = 'nice-scaling-axis mb-3 p-3 text-light';
              axisNode.setAttribute('data-axis-index', index);
              const extremes = axis && typeof axis.getExtremes === 'function' ? axis.getExtremes() : { min: axis.min, max: axis.max };
              const currentMin = Number.isFinite(extremes.min) ? extremes.min : 0;
              const currentMax = Number.isFinite(extremes.max) ? extremes.max : 0;
              const contextStat = context && context.stat ? context.stat : {};
              const formatContext = { isPercent: !!contextStat.isPercent, isMoney: !!contextStat.isMoney };
              const currentLabel = `${formatNumberForDisplay(currentMin, 0, window.Highcharts, formatContext)} – ${formatNumberForDisplay(currentMax, 0, window.Highcharts, formatContext)}`;
              axisNode.innerHTML = `
                <div class="nice-scaling-axis-header d-flex justify-content-between align-items-start mb-3">
                  <div>
                    <div class="axis-label text-uppercase fw-semibold small">${axisLabel}</div>
                    <div class="axis-range mt-1">Current range: ${currentLabel}</div>
                  </div>
                  <div class="badge bg-primary bg-opacity-25 text-light border border-primary border-opacity-30">${index === 0 ? 'Primary' : 'Linked'}</div>
                </div>
                <div class="row g-3">
                  <div class="col-6">
                    <label class="form-label form-label-sm text-uppercase text-light text-opacity-75">Minimum</label>
                    <input type="number" step="any" class="form-control form-control-sm" data-role="min" value="${Number.isFinite(currentMin) ? currentMin.toFixed(3) : ''}" placeholder="Auto" />
                  </div>
                  <div class="col-6">
                    <label class="form-label form-label-sm text-uppercase text-light text-opacity-75">Maximum</label>
                    <input type="number" step="any" class="form-control form-control-sm" data-role="max" value="${Number.isFinite(currentMax) ? currentMax.toFixed(3) : ''}" placeholder="Auto" />
                  </div>
                </div>`;
              const override = overrides[index];
              if (override) {
                const minInput = axisNode.querySelector('input[data-role="min"]');
                const maxInput = axisNode.querySelector('input[data-role="max"]');
                if (minInput) minInput.value = override.min;
                if (maxInput) maxInput.value = override.max;
              }
              scalingModalAxesContainer.appendChild(axisNode);
            });
            scalingModalInstance.show();
          };

          const ensureContainerOverlay = (container) => {
            if (!container) return null;
            const computed = window.getComputedStyle ? window.getComputedStyle(container) : null;
            if (computed && computed.position === 'static') {
              container.style.position = 'relative';
            }
            let overlay = container.querySelector('.nice-scaling-overlay');
            if (!overlay) {
              overlay = document.createElement('div');
              overlay.className = 'nice-scaling-overlay';
              overlay.style.position = 'absolute';
              overlay.style.top = '6px';
              overlay.style.right = '6px';
              overlay.style.zIndex = '30';
              overlay.style.display = 'flex';
              overlay.style.gap = '6px';
              container.appendChild(overlay);
            }
            return overlay;
          };

          const attachScalingControls = (chart, context) => {
            if (!chart || !context || !context.container) return;
            const overlay = ensureContainerOverlay(context.container);
            if (!overlay) return;
            let button = overlay.querySelector('button[data-role="nice-scale"]');
            if (!button) {
              button = document.createElement('button');
              button.type = 'button';
              button.className = 'btn btn-dark btn-sm chart-scale-btn';
              button.setAttribute('data-role', 'nice-scale');
              button.innerHTML = '<span aria-hidden="true">⚙️</span> Scale';
              overlay.appendChild(button);
            }
            button.style.opacity = '0.85';
            button.style.borderRadius = '6px';
            button.style.padding = '2px 8px';
            button.style.fontSize = '11px';
            button.style.transition = 'all .2s ease';
            button.style.border = '1px solid rgba(255,255,255,0.25)';
            button.onmouseenter = () => { button.style.opacity = '1'; };
            button.onmouseleave = () => { button.style.opacity = '0.85'; };
            button.onclick = () => openScalingModal(chart, context);
          };

          const computeAxisBounds = window.computeAxisBounds || ((stat) => ({
            min: 0,
            max: 1,
            ticks: [0, 0.5, 1],
            decimals: 0,
            reversed: !!stat?.upsideDown
          }));

          function buildTrendSeries(stat, config = {}) {
            const prefix = config.prefix ?? (stat.isMoney ? '$' : '');
            const suffix = config.suffix ?? (stat.isPercent ? '%' : '');
            const decimals = Number.isFinite(config.decimals) ? config.decimals : 0;
            const lineWidth = config.lineWidth ?? 3.6;
            const zoneColors = {
              up: (config.zoneColors && config.zoneColors.up) || TREND_UP,
              down: (config.zoneColors && config.zoneColors.down) || TREND_DOWN
            };
            const formatContext = { isPercent: !!stat.isPercent, isMoney: !!stat.isMoney };

            const markerDefaults = {
              enabled: false,
              radius: 0,
              symbol: 'circle',
              lineWidth: 0,
              fillColor: 'transparent',
              lineColor: 'transparent',
              states: {
                hover: { enabled: false }
              }
            };
            const marker = deepMerge(markerDefaults, config.marker || {});

            const dataLabelDefaults = {
              enabled: false,
              rotation: -90,
              color: '#0f172a',
              backgroundColor: 'rgba(255,255,255,0.82)',
              borderRadius: 3,
              borderWidth: 0,
              padding: 1,
              style: {
                fontWeight: '700',
                fontSize: '12px',
                textOutline: 'none',
                whiteSpace: 'nowrap'
              },
              align: 'center',
              verticalAlign: 'top',
              y: -LABEL_RISER_OFFSET_PX,
              x: 0,
              allowOverlap: false,
              overflow: 'allow',
              crop: false
            };
            const dataLabels = deepMerge(dataLabelDefaults, config.dataLabels || {});
            const dataLabelsEnabled = dataLabels.enabled !== false;

            const Highcharts = window.Highcharts;

            return (stat.series || []).map((series, seriesIndex, allSeries) => {
              const data = Array.isArray(series.data) ? series.data.slice() : [];
              const segments = [];
              for (let i = 1; i < data.length; i += 1) {
                const prev = data[i - 1];
                const curr = data[i];
                if (!Number.isFinite(prev) || !Number.isFinite(curr)) continue;
                const color = curr >= prev ? zoneColors.up : zoneColors.down;
                segments.push({ index: i, color });
              }
              const defaultColor = segments.length ? segments[segments.length - 1].color : zoneColors.up;
              const dashStyles = ['Solid', 'ShortDash', 'ShortDot', 'ShortDashDot'];
              const dashStyle = dashStyles[Math.min(seriesIndex, dashStyles.length - 1)] || 'Solid';
              const zones = [];
              if (segments.length > 1) {
                for (let i = 0; i < segments.length - 1; i += 1) {
                  zones.push({ value: segments[i].index, color: segments[i].color });
                }
              } else if (segments.length === 1) {
                zones.push({ value: segments[0].index, color: segments[0].color });
              }

              const labelFormatter = function formatter() {
                if (this.y == null) return '';
                const chartInstance = this.series && this.series.chart;
                if (chartInstance && chartInstance.series && chartInstance.series.length) {
                  const primarySeries = chartInstance.series[0];
                  if (primarySeries && primarySeries !== this.series && primarySeries.visible !== false) {
                    let primaryPoint = null;
                    if (primarySeries.points && primarySeries.points.length) {
                      const sourcePoint = this.point;
                      if (sourcePoint && sourcePoint.index != null && primarySeries.points[sourcePoint.index] && primarySeries.points[sourcePoint.index].x === this.x) {
                        primaryPoint = primarySeries.points[sourcePoint.index];
                      } else {
                        primaryPoint = primarySeries.points.find((candidate) => candidate && candidate.x === this.x) || null;
                      }
                    }
                    if (primaryPoint && Number.isFinite(primaryPoint.y) && Number.isFinite(this.y) && Math.abs(primaryPoint.y - this.y) < 1e-7) {
                      return '';
                    }
                  }
                }
                if (!Highcharts) {
                  const rendered = decimals > 0 ? Number(this.y.toFixed(decimals)) : Math.round(this.y);
                  return `${prefix}${rendered}${suffix}`;
                }
                const dynamicDecimals = Math.max(decimals, Math.min(4, countDecimals(this.y ?? 0)));
                const rendered = formatNumberForDisplay(this.y, dynamicDecimals, Highcharts, formatContext);
                return prefix + rendered + suffix;
              };

              const seriesConfig = {
                ...series,
                data,
                color: defaultColor,
                lineColor: defaultColor,
                lineWidth,
                zoneAxis: 'x',
                zones,
                dashStyle,
                marker
              };

              if (dataLabelsEnabled) {
                seriesConfig.dataLabels = {
                  ...dataLabels,
                  style: {
                    ...(dataLabels.style || {})
                  },
                  formatter: labelFormatter,
                  y: dataLabels.y != null ? dataLabels.y : dataLabelDefaults.y,
                  x: dataLabels.x != null ? dataLabels.x : dataLabelDefaults.x
                };
              } else {
                seriesConfig.dataLabels = { enabled: false };
              }

              return seriesConfig;
            });
          }

          function renderLineChart(config) {
            const {
              container,
              stat,
              height = 520,
              categories = [],
              numberFormat = {},
              chart = {},
              xAxis = {},
              yAxis = {},
              dataLabels = {},
              marker = {},
              lineWidth,
              zoneColors,
              tooltip = {},
              plotOptions = {},
              axis = {}
            } = config || {};

            if (!container || !window.Highcharts) return null;

            let dataPrecision = 0;
            let maxAbsValue = 0;
            let minAbsValue = Number.POSITIVE_INFINITY;
            if (stat && Array.isArray(stat.series)) {
              stat.series.forEach((series) => {
                (series.data || []).forEach((value) => {
                  if (!Number.isFinite(value)) return;
                  dataPrecision = Math.max(dataPrecision, countDecimals(value));
                  const absVal = Math.abs(value);
                  maxAbsValue = Math.max(maxAbsValue, absVal);
                  if (absVal > 0) {
                    minAbsValue = Math.min(minAbsValue, absVal);
                  }
                });
              });
            }
            if (!Number.isFinite(minAbsValue)) {
              minAbsValue = 0;
            }
            if (Number.isFinite(stat?.min)) {
              maxAbsValue = Math.max(maxAbsValue, Math.abs(Number(stat.min)));
            }
            if (Number.isFinite(stat?.max)) {
              maxAbsValue = Math.max(maxAbsValue, Math.abs(Number(stat.max)));
            }

            const axisOptions = { ...axis };
            const desiredTickCountOption = Number(axisOptions.desiredTickCount);
            const desiredTickCountEstimate = Number.isFinite(desiredTickCountOption)
              ? desiredTickCountOption
              : (Number(axisOptions.desiredTickCount) || 6);
            if (!Number.isFinite(axisOptions.desiredTickCount)) {
              axisOptions.desiredTickCount = desiredTickCountEstimate;
            }
            if (!Number.isFinite(axisOptions.minimumTickInterval)) {
              const explicitMinValue = Number.isFinite(stat?.min) ? Number(stat.min) : null;
              const explicitMaxValue = Number.isFinite(stat?.max) ? Number(stat.max) : null;
              let inferredPrecision = dataPrecision;
              if (explicitMinValue != null && explicitMaxValue != null && explicitMaxValue > explicitMinValue) {
                const explicitSpan = explicitMaxValue - explicitMinValue;
                const segmentCount = Math.max(1, (Number.isFinite(axisOptions.desiredTickCount) ? Number(axisOptions.desiredTickCount) : desiredTickCountEstimate) - 1);
                if (segmentCount > 0) {
                  const segmentSize = explicitSpan / segmentCount;
                  if (Number.isFinite(segmentSize)) {
                    inferredPrecision = Math.max(inferredPrecision, countDecimals(segmentSize));
                  }
                }
              }
              const precisionCap = stat && stat.isPercent ? 3 : 4;
              const boundedPrecision = Math.min(precisionCap, Math.max(inferredPrecision, 0));
              if (boundedPrecision > 0) {
                axisOptions.minimumTickInterval = Math.pow(10, -boundedPrecision);
              } else {
                axisOptions.minimumTickInterval = stat && stat.isPercent ? 1 : 1;
              }
            }

            const chartKey = String(config.chartKey || container.dataset.scaleKey || `chart-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`);
            container.dataset.scaleKey = chartKey;

            const prefix = numberFormat.prefix ?? (stat && stat.isMoney ? '$' : '');
            const suffix = numberFormat.suffix ?? (stat && stat.isPercent ? '%' : '');
            const formatContext = { isPercent: !!(stat && stat.isPercent), isMoney: !!(stat && stat.isMoney) };
            const allowMultipleAxes = axisOptions.allowMultipleAxes === true;
            const axisStats = allowMultipleAxes ? deriveAxisStats(stat).slice(0, 4) : [stat];
            const rawAxisBoundsList = axisStats.map((axisStat) => computeAxisBounds(axisStat, axisOptions));
            const axisBoundsList = alignAxisBoundsToPrimary(rawAxisBoundsList);
            const primaryBounds = axisBoundsList[0] || { min: 0, max: 1, ticks: [0, 1], reversed: !!(stat && stat.upsideDown) };
            const axisCount = allowMultipleAxes ? (axisBoundsList.length || 1) : 1;
            const axisMagnitude = Math.max(
              Math.abs(primaryBounds.min ?? 0),
              Math.abs(primaryBounds.max ?? 0),
              maxAbsValue
            );
            let baseDecimals;
            if (stat && stat.isPercent) {
              baseDecimals = axisMagnitude < 1 ? Math.min(Math.max(dataPrecision, 1), 2) : 0;
            } else if (stat && stat.isMoney) {
              baseDecimals = axisMagnitude < 1 ? Math.min(dataPrecision, 2) : Math.min(Math.max(dataPrecision, 1), 1);
            } else if (axisMagnitude >= 1) {
              baseDecimals = Math.min(Math.max(dataPrecision, 1), 1);
            } else {
              baseDecimals = Math.min(Math.max(dataPrecision, 0), 4);
            }
            const primaryDisplayDecimals = Number.isFinite(primaryBounds.displayDecimals) ? primaryBounds.displayDecimals : (primaryBounds.decimals || 0);
            let decimals = Math.max(baseDecimals, primaryDisplayDecimals);
            if (stat && stat.isPercent) {
              decimals = Math.min(decimals, 1);
            } else if (axisMagnitude >= 1) {
              decimals = Math.min(decimals, axisMagnitude < 10 ? 2 : (stat && stat.isMoney ? 1 : 1));
            }
            const trendSeries = buildTrendSeries(stat, {
              prefix,
              suffix,
              decimals,
              dataLabels,
              marker,
              lineWidth,
              zoneColors
            });
            const seriesOrigins = trendSeries.map((_, idx) => Math.min(idx, axisCount - 1));
            trendSeries.forEach((seriesConfig, index) => {
              if (!seriesConfig) return;
              const originAxis = seriesOrigins[index] ?? 0;
              seriesConfig.yAxis = 0;
              seriesConfig.__originAxis = originAxis;
              if (seriesConfig.zIndex == null) {
                seriesConfig.zIndex = index === 0 ? 3 : 2;
              }
            });
            const Highcharts = window.Highcharts;
            const userEvents = (chart && chart.events) || {};
            const mergedEvents = {};
            Object.keys(userEvents).forEach((key) => {
              const handler = userEvents[key];
              if (typeof handler !== 'function') {
                mergedEvents[key] = handler;
                return;
              }
              if (typeof mergedEvents[key] === 'function') {
                const baseHandler = mergedEvents[key];
                mergedEvents[key] = function mergedEvent(...args) {
                  baseHandler.apply(this, args);
                  handler.apply(this, args);
                };
              } else {
                mergedEvents[key] = handler;
              }
            });

            const chartOptions = {
              chart: {
                type: 'line',
                backgroundColor: '#ffffff',
                height,
                spacing: chart.spacing || [24, 18, 24, 24],
                events: mergedEvents
              },
              credits: { enabled: false },
              exporting: { enabled: false },
              title: { text: null },
              xAxis: deepMerge({
                categories,
                lineColor: '#0f172a',
                lineWidth: 2,
                tickColor: '#0f172a',
                tickLength: 11,
                labels: {
                  rotation: -90,
                  align: 'center',
                  x: 0,
                  y: 45,
                  style: { color: '#0f172a', fontSize: '12px', fontWeight: '600' }
                }
              }, xAxis),
              yAxis: (() => {
                const baseAxisConfig = {
                  title: { text: null },
                  lineColor: '#0f172a',
                  lineWidth: 2,
                  tickColor: '#0f172a',
                  tickLength: 11,
                  gridLineColor: 'rgba(15,23,42,0.12)',
                  labels: {
                    style: { color: '#0f172a', fontSize: '12px', fontWeight: '600' },
                    align: 'right',
                    x: -8,
                    formatter() {
                      const axisInstance = this.axis;
                      const bounds = axisInstance && axisInstance.boundsInfo;
                      const displayDecimals = bounds && Number.isFinite(bounds.displayDecimals)
                        ? bounds.displayDecimals
                        : (axisMagnitude < 10 ? 2 : decimals);
                       const rendered = formatNumberForDisplay(this.value, displayDecimals, Highcharts, formatContext);
                      return prefix + rendered + suffix;
                    }
                  }
                };
                if (axisCount <= 1) {
                  const singleAxis = deepMerge({
                    ...baseAxisConfig,
                    min: primaryBounds.min,
                    max: primaryBounds.max,
                    tickPositions: primaryBounds.ticks,
                    startOnTick: true,
                    endOnTick: true,
                    allowDecimals: primaryBounds.allowDecimals !== false,
                    reversed: primaryBounds.reversed
                  }, yAxis);
                  const singleLabels = singleAxis.labels || (singleAxis.labels = {});
                  if (singleLabels.align == null) {
                    singleLabels.align = 'right';
                  }
                  if (singleLabels.x == null) {
                    singleLabels.x = -8;
                  }
                  singleAxis.boundsInfo = primaryBounds;
                  return singleAxis;
                }
                return axisBoundsList.map((bounds, index) => {
                  const isPrimary = index === 0;
                  const displayDecimals = bounds && Number.isFinite(bounds.displayDecimals)
                    ? bounds.displayDecimals
                    : (isPrimary ? decimals : Math.min(decimals, 2));
                  const axisConfig = deepMerge({
                    ...baseAxisConfig,
                    min: bounds?.min,
                    max: bounds?.max,
                    tickPositions: bounds?.ticks,
                    startOnTick: true,
                    endOnTick: true,
                    allowDecimals: bounds?.allowDecimals !== false,
                    reversed: bounds?.reversed,
                    visible: isPrimary,
                    opposite: !isPrimary && index % 2 === 1
                  }, yAxis);
                  axisConfig.labels = axisConfig.labels || {};
                  if (axisConfig.opposite) {
                    if (axisConfig.labels.align == null) {
                      axisConfig.labels.align = 'left';
                    }
                    if (axisConfig.labels.x == null) {
                      axisConfig.labels.x = 8;
                    }
                  } else {
                    if (axisConfig.labels.align == null) {
                      axisConfig.labels.align = 'right';
                    }
                    if (axisConfig.labels.x == null) {
                      axisConfig.labels.x = -8;
                    }
                  }
                  const userFormatter = axisConfig.labels.formatter;
                  axisConfig.labels.formatter = function patchedFormatter() {
                    if (this.axis && axisConfig.boundsInfo) {
                      this.axis.boundsInfo = axisConfig.boundsInfo;
                    }
                    if (typeof userFormatter === 'function') {
                      return userFormatter.call(this);
                    }
                    const rendered = formatNumberForDisplay(this.value, displayDecimals, Highcharts, formatContext);
                    return prefix + rendered + suffix;
                  };
                  axisConfig.userFormatter = userFormatter;
                  axisConfig.boundsInfo = {
                    ...bounds,
                    displayDecimals
                  };
                  if (!isPrimary) {
                    axisConfig.lineWidth = 0;
                    axisConfig.tickLength = 0;
                    axisConfig.gridLineWidth = 0;
                    axisConfig.labels = axisConfig.labels || {};
                    axisConfig.labels.enabled = false;
                  }
                  return axisConfig;
                });
              })(),
              legend: { enabled: false },
              tooltip: deepMerge({
                shared: true,
                useHTML: true,
                formatter() {
                  return '<b>' + escapeHtml(this.x) + '</b>' + this.points
                    .map((pt) => {
                      const dynamicDecimals = Math.max(decimals, Math.min(4, countDecimals(pt.y ?? 0)));
                      const val = pt.y == null
                        ? '—'
                        : prefix + formatNumberForDisplay(pt.y, dynamicDecimals, Highcharts, formatContext) + suffix;
                      return `<br><span style="color:${pt.color}">●</span> ${escapeHtml(pt.series.name)}: <b>${val}</b>`;
                    })
                    .join('');
                }
              }, tooltip),
              plotOptions: deepMerge({
                series: {
                  marker: deepMerge({ enabled: false, radius: 0, lineWidth: 0, symbol: 'circle' }, marker),
                  lineWidth: lineWidth || 3.6,
                  connectNulls: false,
                  dataLabels: { enabled: true }
                }
              }, plotOptions),
              series: trendSeries
            };

            ensureSoftWhiteHaloFilter();
            ensureChartDecorators();

            const chartInstance = Highcharts.chart(container, chartOptions);
            applyDataLabelStyles(chartInstance);
            installDroplines(chartInstance);

            const context = {
              container,
              stat,
              axisOptions,
              chartKey,
              axisStats,
              autoBounds: axisBoundsList,
              seriesOrigins
            };
            chartInstance.__scaleContext = context;
            chartInstance.__seriesOrigins = seriesOrigins.slice();
            chartInstance.__rebindSeries = (overrides) => rebindChartSeries(chartInstance, context, overrides);

            applyAutoScaling(chartInstance, context);
            rebindChartSeries(chartInstance, context);
            chartInstance.redraw();

            const storedOverrides = fetchManualScaleOverrides(chartKey);
            if (Object.keys(storedOverrides).length) {
              applyManualOverrides(chartInstance, storedOverrides, context);
            }

            attachScalingControls(chartInstance, context);

            return chartInstance;
          }

          return {
            computeAxisBounds,
            buildTrendSeries,
            renderLineChart,
            attachScalingControls,
            applyManualOverrides,
            applyAutoScaling,
            constants: measurementConstants
          };
        })();
      };

      const childRenderSource = () => {
        const ensureScaleHelpers = () => {
          if (window.__BOOKLET_SCALE_KEY_HELPERS__) {
            return window.__BOOKLET_SCALE_KEY_HELPERS__;
          }
          const normalizeManualScaleKeyPart = (value, fallback) => {
            const source = value == null || value === '' ? fallback : value;
            if (source == null || source === '') return '';
            let text = String(source);
            if (typeof text.normalize === 'function') {
              text = text.normalize('NFKD').replace(/[\u0300-\u036F]/g, '');
            }
            return text
              .trim()
              .toLowerCase()
              .replace(/[:|]+/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();
          };

          const ensureStatManualScaleKey = (division, staff, stat, indices = {}) => {
            if (!stat || typeof stat !== 'object') {
              const divisionIndex = Number.isFinite(indices.division) ? indices.division : 0;
              const staffIndex = Number.isFinite(indices.staff) ? indices.staff : 0;
              const statIndex = Number.isFinite(indices.stat) ? indices.stat : 0;
              return `chart-${divisionIndex}-${staffIndex}-${statIndex}`;
            }
            if (typeof stat.manualScaleKey === 'string' && stat.manualScaleKey.length) {
              return stat.manualScaleKey;
            }
            const divisionIndex = Number.isFinite(indices.division) ? indices.division : 0;
            const staffIndex = Number.isFinite(indices.staff) ? indices.staff : 0;
            const statIndex = Number.isFinite(indices.stat) ? indices.stat : 0;
            const parts = [];
            const divisionPart = normalizeManualScaleKeyPart(division && division.name, `division-${divisionIndex + 1}`);
            if (divisionPart) parts.push(divisionPart);
            const staffNamePart = normalizeManualScaleKeyPart(staff && staff.staffName, `staff-${staffIndex + 1}`);
            if (staffNamePart) parts.push(staffNamePart);
            const staffPostPart = normalizeManualScaleKeyPart(staff && staff.post);
            if (staffPostPart) parts.push(staffPostPart);
            const statKeyPart = normalizeManualScaleKeyPart(stat && stat.key, `stat-${statIndex + 1}`);
            if (statKeyPart) parts.push(statKeyPart);
            const statLabelPart = normalizeManualScaleKeyPart(stat && stat.displayName);
            if (statLabelPart && statLabelPart !== statKeyPart) {
              parts.push(statLabelPart);
            }
            const manualKey = parts.filter(Boolean).join('::') || `chart-${divisionIndex}-${staffIndex}-${statIndex}`;
            stat.manualScaleKey = manualKey;
            return manualKey;
          };

          const slugifyForId = (value, fallback = 'item') => {
            let text = value == null ? '' : String(value);
            if (typeof text.normalize === 'function') {
              text = text.normalize('NFKD').replace(/[\u0300-\u036F]/g, '');
            }
            const slug = text
              .toLowerCase()
              .replace(/[^a-z0-9]+/g, '-')
              .replace(/-+/g, '-')
              .replace(/^-|-$/g, '');
            return slug || fallback;
          };

          const helpers = { ensureStatManualScaleKey, slugifyForId };
          window.__BOOKLET_SCALE_KEY_HELPERS__ = helpers;
          return helpers;
        };

        const { ensureStatManualScaleKey, slugifyForId } = ensureScaleHelpers();
        const data = window.__BOOKLET_DATA || { divisions: [] };
        const printBtn = document.getElementById('printBtn');
        if (printBtn) printBtn.addEventListener('click', () => window.print());
        const wrapper = document.getElementById('pages');
        if (!wrapper) return;

        const pages = [];
        const charts = [];

        if (!window.__BOOKLET_SHARED_RUNTIME__) {
          console.error('Shared chart runtime unavailable.');
          return;
        }
    const sharedRuntime = window.__BOOKLET_SHARED_RUNTIME__;
  const axisDefaults = { desiredTickCount: 6, padFraction: 0.08, clampZero: true, minimumTickInterval: 5, allowMultipleAxes: true };
    const measurement = (sharedRuntime && sharedRuntime.constants) || null;
    const DEFAULT_CM_TO_PX = 37.7952755906;
    const fallbackTick = Math.round(0.5 * DEFAULT_CM_TO_PX);
    const fallbackClearance = Math.round(0.3 * DEFAULT_CM_TO_PX);
    const MARKER_TICK_LENGTH_PX = Number.isFinite(measurement?.MARKER_TICK_LENGTH_PX)
      ? measurement.MARKER_TICK_LENGTH_PX
      : fallbackTick;
    const LABEL_CLEARANCE_PX = Number.isFinite(measurement?.LABEL_CLEARANCE_PX)
      ? measurement.LABEL_CLEARANCE_PX
      : fallbackClearance;
    const LABEL_LIFT_PX = Number.isFinite(measurement?.LABEL_RISER_OFFSET_PX)
      ? measurement.LABEL_RISER_OFFSET_PX
      : MARKER_TICK_LENGTH_PX + LABEL_CLEARANCE_PX;

        function escapeHtml(str) {
          return String(str || '').replace(/[&<>"]|'/g, (c) => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
          }[c]));
        }

        function textToHtml(str) {
          return escapeHtml(str).replace(/\n/g, '<br>');
        }

        function splitVfpText(raw) {
          const normalized = String(raw || '').replace(/\r\n?/g, '\n');
          const parts = normalized
            .split(/\n{2,}/)
            .map((segment) => segment.trim())
            .filter((segment) => segment.length);
          return parts.length ? parts : ['VFP'];
        }

        function formatRange() {
          if (!data.earliest || !data.latest) return '';
          const start = new Date(data.earliest);
          const end = new Date(data.latest);
          if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime())) return '';
          const fmt = (d) => {
            const day = d.getDate();
            const mon = d.toLocaleDateString(undefined, { month: 'short' });
            const yr = d.toLocaleDateString(undefined, { year: '2-digit' });
            return `${day} ${mon} ${yr}`;
          };
          const startLabel = fmt(start);
          const endLabel = fmt(end);
          return startLabel === endLabel ? endLabel : `${startLabel} – ${endLabel}`;
        }

        function formatMeta(staff) {
          const parts = [];
          if (staff.post) parts.push(staff.post);
          if (staff.staffName) parts.push(staff.staffName);
          return parts.map(escapeHtml).join('<br>');
        }

        const rangeLabel = formatRange();

        const renderCoverPage = ({ title, subtitle, range, logo }) => `
          <section class="page page-cover" style="--base: #ffffff;">
            <div class="page-inner">
              <h1>${escapeHtml(title)}</h1>
              ${subtitle ? `<h2>${escapeHtml(subtitle)}</h2>` : ''}
              ${range ? `<div class="text-uppercase" style="letter-spacing:0.24em;color:rgba(15,23,42,0.6);font-weight:700;">${escapeHtml(range)}</div>` : ''}
              ${logo ? `<img src="${logo}" alt="Logo" />` : ''}
            </div>
          </section>`;

        const renderDivisionLogoPage = ({ baseColor, logo }) => `
          <section class="page page-logo" style="--base:${baseColor};">
            <div class="page-inner" style="background: radial-gradient(circle at center, rgba(255,255,255,0.92) 0%, ${baseColor} 72%);">
              <img src="${logo}" alt="Division Logo" />
            </div>
          </section>`;

        const renderVfpPage = ({ division, classes, baseColor, textColor, accentColor, encodedParagraphs, initialParagraph }) => `
          <section class="page page-vfp${classes}" style="--base:${baseColor};">
            <div class="page-inner" style="color:${textColor};">
              <div class="vfp-body">
                <div class="vfp-heading">
                  <h2>${escapeHtml(division.name)} DIVISION</h2>
                  <div class="vfp-label" style="color:${accentColor};">VFP</div>
                </div>
                <div class="vfp-text" data-paragraphs="${encodedParagraphs}">${initialParagraph}</div>
              </div>
            </div>
          </section>`;

        const renderStatPage = ({ division, staff, stat, chartId }) => `
          <section class="page page-stat" style="--base:${division.color.base};">
            <div class="page-inner">
              <div class="info">
                <h3>${escapeHtml(stat.title || stat.displayName)}</h3>
                <div class="meta">${formatMeta(staff) || '&nbsp;'}</div>
              </div>
              <div class="divider"></div>
              <div class="chart" id="${chartId}"></div>
            </div>
          </section>`;

        const renderClosingLogoPage = (logo) => `
          <section class="page page-logo" style="--base:#ffffff;">
            <div class="page-inner" style="background: radial-gradient(circle at center, rgba(255,255,255,0.95) 0%, #ffffff 70%);">
              <img src="${logo}" alt="Logo" />
            </div>
          </section>`;

        pages.push(renderCoverPage({
          title: data.title || 'Weekly BP',
          subtitle: data.subtitle || '',
          range: rangeLabel,
          logo: data.logo || ''
        }));

        data.divisions.forEach((division, di) => {
          if (data.logo) {
            const logoBase = division.color && division.color.base ? division.color.base : '#ffffff';
            pages.push(renderDivisionLogoPage({ baseColor: logoBase, logo: data.logo }));
          }

          const vfpScale = division.vfpScale === 'compact' || division.vfpScale === 'large' ? division.vfpScale : 'normal';
          const vfpAlign = ['left', 'right', 'center'].includes(division.vfpAlign) ? division.vfpAlign : 'center';
          const vfpWidth = ['standard', 'wide', 'narrow'].includes(division.vfpWidth) ? division.vfpWidth : 'standard';
          const vfpFont = ['small', 'large', 'xlarge'].includes(division.vfpFont) ? division.vfpFont : 'medium';
          const scaleClass = vfpScale === 'normal' ? '' : ` scale-${vfpScale}`;
          const alignClass = vfpAlign === 'center' ? '' : ` align-${vfpAlign}`;
          const widthClass = vfpWidth === 'standard' ? '' : ` width-${vfpWidth}`;
          const fontClass = vfpFont === 'medium' ? '' : ` font-${vfpFont}`;
          const accentColor = division.color && division.color.accent ? division.color.accent : '#334155';
          const baseColor = division.color && division.color.base ? division.color.base : '#f8fafc';
          const textColor = division.color && division.color.text ? division.color.text : '#0F172A';
          const paragraphs = splitVfpText(division.vfpText || 'VFP');
          const paragraphHtml = paragraphs.map((paragraph) => textToHtml(paragraph));
          const encodedParagraphs = encodeURIComponent(JSON.stringify(paragraphHtml));
          const initialParagraph = paragraphHtml[0] || textToHtml('VFP');
          pages.push(renderVfpPage({
            division,
            classes: `${scaleClass}${alignClass}${widthClass}${fontClass}`,
            baseColor,
            textColor,
            accentColor,
            encodedParagraphs,
            initialParagraph
          }));

          division.staff.forEach((staff, si) => {
            staff.stats.forEach((stat, ti) => {
              const manualScaleKey = ensureStatManualScaleKey(division, staff, stat, { division: di, staff: si, stat: ti });
              const fallbackId = `${di}-${si}-${ti}`;
              const slugSource = `${manualScaleKey} ${fallbackId}`;
              const chartId = `chart-${slugifyForId(slugSource, fallbackId)}`;
              charts.push({ id: chartId, stat, chartKey: manualScaleKey });
              pages.push(renderStatPage({ division, staff, stat, chartId }));
            });
          });
        });

        if (data.logo) {
          pages.push(renderClosingLogoPage(data.logo));
        }

        wrapper.innerHTML = pages.join('');

        function layoutVfpPages() {
          const sections = Array.from(wrapper.querySelectorAll('.page-vfp'));
          sections.forEach((section) => {
            const textHolder = section.querySelector('.vfp-text');
            if (!textHolder) return;
            const dataAttr = textHolder.getAttribute('data-paragraphs');
            if (!dataAttr) return;
            let paragraphs;
            try {
              paragraphs = JSON.parse(decodeURIComponent(dataAttr));
            } catch (err) {
              paragraphs = [textHolder.innerHTML || ''];
            }
            if (!Array.isArray(paragraphs) || !paragraphs.length) {
              paragraphs = [textHolder.innerHTML || ''];
            }
            textHolder.removeAttribute('data-paragraphs');
            const template = section.cloneNode(true);
            const templateHolder = template.querySelector('.vfp-text');
            if (templateHolder) {
              templateHolder.innerHTML = '';
              templateHolder.removeAttribute('data-paragraphs');
            }
            const templateHeading = template.querySelector('.vfp-heading');
            const continuationTemplate = template.cloneNode(true);
            const continuationHolder = continuationTemplate.querySelector('.vfp-text');
            if (continuationHolder) {
              continuationHolder.innerHTML = '';
            }
            const continuationHeading = continuationTemplate.querySelector('.vfp-heading');
            if (continuationHeading && continuationHeading.parentNode) {
              continuationHeading.parentNode.removeChild(continuationHeading);
            }
            const templateInner = template.querySelector('.page-inner');
            const baseHolder = textHolder;
            const baseInner = section.querySelector('.page-inner');
            baseHolder.innerHTML = '';

            let currentSection = section;
            let currentHolder = baseHolder;
            let currentInner = baseInner;

            function appendToNewPage(paragraphEl) {
              const newSection = continuationTemplate.cloneNode(true);
              const newHolder = newSection.querySelector('.vfp-text');
              const newInner = newSection.querySelector('.page-inner');
              if (newHolder) {
                newHolder.innerHTML = '';
                newHolder.removeAttribute('data-paragraphs');
                newHolder.appendChild(paragraphEl);
              }
              const parent = currentSection.parentNode;
              if (parent) {
                parent.insertBefore(newSection, currentSection.nextSibling);
              }
              currentSection = newSection;
              currentHolder = newHolder;
              currentInner = newInner;
              currentSection.classList.add('page-vfp-continuation');
            }

            paragraphs.forEach((html) => {
              const paragraphEl = document.createElement('p');
              paragraphEl.innerHTML = html;
              currentHolder.appendChild(paragraphEl);
              if (currentInner && currentInner.scrollHeight > currentInner.clientHeight + 2) {
                currentHolder.removeChild(paragraphEl);
                appendToNewPage(paragraphEl);
              }
            });
          });
        }

        layoutVfpPages();

        function renderChart(cfg) {
          const { id, stat, chartKey } = cfg || {};
          const container = document.getElementById(id);
          if (!container) return;
          const inner = container.closest('.page-inner');
          const available = inner ? inner.clientHeight : container.parentElement ? container.parentElement.clientHeight : 640;
          const chartHeight = Math.max(available - 60, 520);
          container.style.height = chartHeight + 'px';
          container.style.width = '100%';
          sharedRuntime.renderLineChart({
            container,
            stat,
            chartKey,
            height: chartHeight,
            categories: stat.dates,
            numberFormat: {
              prefix: stat.isMoney ? '$' : '',
              suffix: stat.isPercent ? '%' : ''
            },
            chart: { spacing: [24, 18, 24, 24] },
            xAxis: {
              labels: {
                rotation: -90,
                align: 'center',
                x: 0,
                y: 45
              }
            },
            dataLabels: {
              verticalAlign: 'top',
              y: -LABEL_LIFT_PX
            },
            marker: { radius: 3, lineWidth: 1.2 },
            lineWidth: 3.6,
            axis: { ...axisDefaults }
          });
        }

        charts.forEach(renderChart);
      };

      const childGraphRenderSource = () => {
        const ensureScaleHelpers = () => {
          if (window.__BOOKLET_SCALE_KEY_HELPERS__) {
            return window.__BOOKLET_SCALE_KEY_HELPERS__;
          }
          const normalizeManualScaleKeyPart = (value, fallback) => {
            const source = value == null || value === '' ? fallback : value;
            if (source == null || source === '') return '';
            let text = String(source);
            if (typeof text.normalize === 'function') {
              text = text.normalize('NFKD').replace(/[\u0300-\u036F]/g, '');
            }
            return text
              .trim()
              .toLowerCase()
              .replace(/[:|]+/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();
          };

          const ensureStatManualScaleKey = (division, staff, stat, indices = {}) => {
            if (!stat || typeof stat !== 'object') {
              const divisionIndex = Number.isFinite(indices.division) ? indices.division : 0;
              const staffIndex = Number.isFinite(indices.staff) ? indices.staff : 0;
              const statIndex = Number.isFinite(indices.stat) ? indices.stat : 0;
              return `chart-${divisionIndex}-${staffIndex}-${statIndex}`;
            }
            if (typeof stat.manualScaleKey === 'string' && stat.manualScaleKey.length) {
              return stat.manualScaleKey;
            }
            const divisionIndex = Number.isFinite(indices.division) ? indices.division : 0;
            const staffIndex = Number.isFinite(indices.staff) ? indices.staff : 0;
            const statIndex = Number.isFinite(indices.stat) ? indices.stat : 0;
            const parts = [];
            const divisionPart = normalizeManualScaleKeyPart(division && division.name, `division-${divisionIndex + 1}`);
            if (divisionPart) parts.push(divisionPart);
            const staffNamePart = normalizeManualScaleKeyPart(staff && staff.staffName, `staff-${staffIndex + 1}`);
            if (staffNamePart) parts.push(staffNamePart);
            const staffPostPart = normalizeManualScaleKeyPart(staff && staff.post);
            if (staffPostPart) parts.push(staffPostPart);
            const statKeyPart = normalizeManualScaleKeyPart(stat && stat.key, `stat-${statIndex + 1}`);
            if (statKeyPart) parts.push(statKeyPart);
            const statLabelPart = normalizeManualScaleKeyPart(stat && stat.displayName);
            if (statLabelPart && statLabelPart !== statKeyPart) {
              parts.push(statLabelPart);
            }
            const manualKey = parts.filter(Boolean).join('::') || `chart-${divisionIndex}-${staffIndex}-${statIndex}`;
            stat.manualScaleKey = manualKey;
            return manualKey;
          };

          const slugifyForId = (value, fallback = 'item') => {
            let text = value == null ? '' : String(value);
            if (typeof text.normalize === 'function') {
              text = text.normalize('NFKD').replace(/[\u0300-\u036F]/g, '');
            }
            const slug = text
              .toLowerCase()
              .replace(/[^a-z0-9]+/g, '-')
              .replace(/-+/g, '-')
              .replace(/^-|-$/g, '');
            return slug || fallback;
          };

          const helpers = { ensureStatManualScaleKey, slugifyForId };
          window.__BOOKLET_SCALE_KEY_HELPERS__ = helpers;
          return helpers;
        };

        const { ensureStatManualScaleKey, slugifyForId } = ensureScaleHelpers();
        const data = window.__BOOKLET_DATA || { divisions: [] };
        const printBtn = document.getElementById('printBtn');
        if (printBtn) printBtn.addEventListener('click', () => window.print());
        const wrapper = document.getElementById('pages');
        if (!wrapper) return;

        const headingMode = data.headingMode || 'stat';
        const layoutMode = data.layoutMode || 'top';
        const showLogo = !!(data.logo && data.logoMode === 'show');
        const pages = [];
        const charts = [];

        if (!window.__BOOKLET_SHARED_RUNTIME__) {
          console.error('Shared chart runtime unavailable.');
          return;
        }
  const sharedRuntime = window.__BOOKLET_SHARED_RUNTIME__;
  const axisDefaults = { desiredTickCount: 6, padFraction: 0.06, clampZero: true, minimumTickInterval: 4, allowMultipleAxes: true };
  const measurement = (sharedRuntime && sharedRuntime.constants) || null;
  const DEFAULT_CM_TO_PX = 37.7952755906;
  const fallbackTick = Math.round(0.5 * DEFAULT_CM_TO_PX);
  const fallbackClearance = Math.round(0.3 * DEFAULT_CM_TO_PX);
  const MARKER_TICK_LENGTH_PX = Number.isFinite(measurement?.MARKER_TICK_LENGTH_PX)
    ? measurement.MARKER_TICK_LENGTH_PX
    : fallbackTick;
  const LABEL_CLEARANCE_PX = Number.isFinite(measurement?.LABEL_CLEARANCE_PX)
    ? measurement.LABEL_CLEARANCE_PX
    : fallbackClearance;
  const LABEL_LIFT_PX = Number.isFinite(measurement?.LABEL_RISER_OFFSET_PX)
    ? measurement.LABEL_RISER_OFFSET_PX
    : MARKER_TICK_LENGTH_PX + LABEL_CLEARANCE_PX;

        function escapeHtml(str) {
          return String(str || '').replace(/[&<>"']|`/g, (c) => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;',
            '`': '&#96;'
          }[c]));
        }

        const metaLabel = data.subtitle ? String(data.subtitle).trim() : '';

        function renderGraphPage({ layoutMode: mode, showLogo: allowLogo, divisionLabel, primaryTitle, secondaryTitle, metaLabelText, chartId, logo }) {
          const safeDivision = (divisionLabel || '').trim();
          const safePrimary = (primaryTitle || '').trim();
          const safeSecondary = (secondaryTitle || '').trim();
          const safeMeta = (metaLabelText || '').trim();
          const overlineMarkup = safeDivision ? `<div class="page-graph-overline">${escapeHtml(safeDivision)}</div>` : '';
          const primaryMarkup = safePrimary ? `<h3 class="page-graph-title">${escapeHtml(safePrimary)}</h3>` : '';
          const secondaryMarkup = safeSecondary ? `<div class="page-graph-subtitle">${escapeHtml(safeSecondary)}</div>` : '';
          const metaMarkup = safeMeta ? `<div class="page-graph-meta">${escapeHtml(safeMeta)}</div>` : '';

          if (mode === 'side') {
            const logoBlock = allowLogo && logo
              ? `<div class="page-graph-logo side"><img src="${logo}" alt="Logo" /></div>`
              : '';
            return `
                  <section class="page page-graph page-graph-side">
                    <div class="page-graph-inner split">
                      <div class="page-graph-info">
                        ${logoBlock}
                        ${overlineMarkup}
                        ${primaryMarkup}
                        ${secondaryMarkup}
                        ${metaMarkup}
                      </div>
                      <div class="page-graph-divider"></div>
                      <div class="page-graph-chart">
                        <div class="chart" id="${chartId}"></div>
                      </div>
                    </div>
                  </section>`;
          }

          const headerClass = allowLogo && logo ? 'page-graph-header simple with-logo' : 'page-graph-header simple';
          const logoBlock = allowLogo && logo
            ? `<div class="page-graph-logo top"><img src="${logo}" alt="Logo" /></div>`
            : '';
          return `
                  <section class="page page-graph">
                    <div class="page-graph-inner">
                      <div class="${headerClass}">
                        ${logoBlock}
                        <div class="page-graph-titles">
                          ${overlineMarkup}
                          ${primaryMarkup}
                          ${secondaryMarkup}
                          ${metaMarkup}
                        </div>
                      </div>
                      <div class="page-graph-chart">
                        <div class="chart" id="${chartId}"></div>
                      </div>
                    </div>
                  </section>`;
        }

        data.divisions.forEach((division, di) => {
          division.staff.forEach((staff, si) => {
            staff.stats.forEach((stat, ti) => {
              const manualScaleKey = ensureStatManualScaleKey(division, staff, stat, { division: di, staff: si, stat: ti });
              const fallbackId = `${di}-${si}-${ti}`;
              const slugSource = `${manualScaleKey} ${fallbackId}`;
              const chartId = `graph-chart-${slugifyForId(slugSource, fallbackId)}`;
              charts.push({ id: chartId, stat, chartKey: manualScaleKey });

              const staffName = (staff.staffName || '').trim();
              const staffPost = (staff.post || '').trim();
              const statLabel = (stat.title || stat.displayName || '').trim();

              const primaryTitle = headingMode === 'staff'
                ? (staffName || staffPost || '')
                : (statLabel || 'Stat');
              const secondaryTitle = headingMode === 'staff'
                ? (statLabel || '')
                : [staffPost, staffName].filter(Boolean).join(' • ');

              const divisionLabel = division.synthetic
                ? (division.name && division.name.trim() ? division.name.trim() : '')
                : (division.name && division.name.trim() ? division.name.trim() : 'Division');

              pages.push(renderGraphPage({
                layoutMode,
                showLogo,
                divisionLabel,
                primaryTitle,
                secondaryTitle,
                metaLabelText: metaLabel,
                chartId,
                logo: data.logo
              }));
            });
          });
        });

        wrapper.innerHTML = pages.join('');

        function renderChart(cfg) {
          const { id, stat, chartKey } = cfg || {};
          const container = document.getElementById(id);
          if (!container) return;
          const holder = container.parentElement;
          const holderHeight = holder ? holder.clientHeight : 600;
          const reserve = Math.max(Math.min(holderHeight * 0.01, 6), 2);
          let chartHeight = holderHeight - reserve;
          chartHeight = Math.max(chartHeight, 440);
          if (chartHeight > holderHeight - 6) {
            chartHeight = Math.max(holderHeight - 6, 360);
          }
          const resolvedHeight = Math.round(chartHeight);
          container.style.height = resolvedHeight + 'px';
          container.style.width = '100%';
          container.style.marginTop = '4px';

          sharedRuntime.renderLineChart({
            container,
            stat,
            chartKey,
            height: resolvedHeight,
            categories: stat.dates,
            numberFormat: {
              prefix: stat.isMoney ? '$' : '',
              suffix: stat.isPercent ? '%' : ''
            },
            chart: { spacing: [12, 12, 18, 12] },
            xAxis: {
              labels: {
                rotation: -90,
                align: 'center',
                x: 0,
                y: 46
              }
            },
            dataLabels: {
              verticalAlign: 'top',
              y: -LABEL_LIFT_PX
            },
            marker: { radius: 3, lineWidth: 1.2 },
            lineWidth: 3.3,
            axis: { ...axisDefaults }
          });
        }

        charts.forEach(renderChart);
      };

      function openBooklet(payload) {
        const safeTitle = escapeHtml(payload.title || 'Weekly BP');
    const pageInfo = resolvePageSizeInfo(payload.pageSize);
    const pageWidth = pageInfo.width;
    const pageHeight = pageInfo.height;
    const pageSizeCss = `${pageWidth} ${pageHeight}`;
  const pagePadding = resolvePagePadding(pageInfo.key);
  const pagePaddingCss = formatPadding(pagePadding);
        const bookletHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    :root { --page-width: ${pageWidth}; --page-height: ${pageHeight}; }
    @page { size: ${pageSizeCss}; margin: 0; }
    *, *::before, *::after { box-sizing: border-box; }
  body { margin: 0; background: #0f172a; font-family: 'Segoe UI', Roboto, Arial, sans-serif; display: flex; flex-direction: column; align-items: center; min-height: 100vh; }
    .toolbar { position: sticky; top: 0; z-index: 10; display: flex; justify-content: flex-end; gap: 12px; padding: 12px 16px; background: rgba(255,255,255,0.92); border-bottom: 1px solid rgba(15,23,42,0.12); box-shadow: 0 6px 20px -12px rgba(15,23,42,0.4); }
    .toolbar button { border: 1px solid rgba(15,23,42,0.2); border-radius: 999px; padding: 6px 16px; background: #0f172a; color: #fff; font-weight: 600; cursor: pointer; }
  .wrapper { width: 100%; max-width: calc(var(--page-width) + 2in); margin: 0 auto; padding: 36px 24px; display: flex; flex-direction: column; gap: 24px; align-items: center; }
  .page { width: var(--page-width); height: var(--page-height); background: #fff; border-radius: 18px; box-shadow: 0 30px 70px -28px rgba(15,23,42,0.55); padding: ${pagePaddingCss}; display: flex; align-items: stretch; position: relative; overflow: hidden; break-after: page; }
  .page-inner { height: 100%; }
  .page:last-child { break-after: auto; }
    .page-inner { flex: 1; background: linear-gradient(135deg, var(--base, #f8fafc) 0%, rgba(255,255,255,0.92) 64%); border: 2px solid rgba(15,23,42,0.12); border-radius: 16px; padding: 28px; display: flex; flex-direction: column; justify-content: center; align-items: center; color: #0f172a; gap: 24px; }
    .page-cover .page-inner { text-align: center; gap: 20px; }
    .page-cover h1 { margin: 0; font-size: 46px; letter-spacing: 0.12em; text-transform: uppercase; font-weight: 800; }
    .page-cover h2 { margin: 0; font-size: 20px; letter-spacing: 0.28em; text-transform: uppercase; color: rgba(15,23,42,0.65); }
    .page-cover img { max-width: 58%; max-height: 58%; object-fit: contain; filter: drop-shadow(0 12px 24px rgba(15,23,42,0.2)); }
    .page-logo { justify-content: center; align-items: center; }
    .page-logo .page-inner { background: radial-gradient(circle at center, rgba(255,255,255,0.92) 0%, var(--base, #ffffff) 70%); }
    .page-logo img { max-width: 65%; max-height: 65%; object-fit: contain; filter: drop-shadow(0 8px 18px rgba(15,23,42,0.25)); }
  .page-vfp .page-inner { padding: 0.7in 0.85in; background: var(--base, #f8fafc); align-items: center; justify-content: flex-start; }
  .page-vfp .vfp-body { width: 100%; max-width: calc(100% - 1.4in); margin: 0 auto; display: flex; flex-direction: column; align-items: center; gap: 26px; text-align: center; }
  .page-vfp.width-wide .vfp-body { max-width: calc(100% - 0.6in); }
  .page-vfp.width-narrow .vfp-body { max-width: calc(100% - 2.3in); }
  .page-vfp.align-left .vfp-body { align-items: flex-start; text-align: left; margin-left: 0.3in; margin-right: auto; }
  .page-vfp.align-right .vfp-body { align-items: flex-end; text-align: right; margin-right: 0.3in; margin-left: auto; }
  .page-vfp.align-left .vfp-text { text-align: left; }
  .page-vfp.align-right .vfp-text { text-align: right; }
  .page-vfp .vfp-heading { width: 100%; display: flex; flex-direction: column; align-items: center; gap: 12px; text-transform: uppercase; letter-spacing: 0.28em; text-align: center; }
  .page-vfp .vfp-heading h2 { margin: 0; font-size: 34px; font-weight: 800; letter-spacing: 0.24em; }
  .page-vfp .vfp-label { font-size: 1.15rem; letter-spacing: 0.42em; font-weight: 800; color: rgba(15,23,42,0.6); }
  .page-vfp .vfp-text { width: 100%; font-size: 20px; font-weight: 700; text-transform: uppercase; color: rgba(15,23,42,0.9); }
  .page-vfp .vfp-text p { margin: 0 0 18px; line-height: 1.6; }
  .page-vfp .vfp-text p:last-child { margin-bottom: 0; }
  .page-vfp.scale-compact .vfp-text p { line-height: 1.45; }
  .page-vfp.scale-large .vfp-text p { line-height: 1.75; }
  .page-vfp.font-small .vfp-text { font-size: 18px; }
  .page-vfp.font-large .vfp-text { font-size: 22px; }
  .page-vfp.font-xlarge .vfp-text { font-size: 24px; }
  .page-vfp-continuation .vfp-body { margin-top: 0; }
  .page-vfp-continuation .page-inner { justify-content: flex-start; }
  .page-stat .page-inner { flex-direction: row; align-items: stretch; gap: 32px; padding: 32px 36px; background: var(--base, #f8fafc); }
  .page-stat .info { width: 2.9in; display: flex; flex-direction: column; gap: 18px; }
  .page-stat .info h3 { margin: 0; font-size: 18px; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 800; line-height: 1.3; }
  .page-stat .info .meta { font-size: 18px; font-weight: 700; text-transform: uppercase; color: rgba(15,23,42,0.85); line-height: 1.35; }
  .page-stat .divider { width: 2px; background: rgba(15,23,42,0.22); border-radius: 999px; }
  .page-stat .chart { flex: 1; min-height: 0; height: 100%; }
    @media print {
      body { background: #fff; margin: 0 !important; min-height: auto !important; display: block !important; }
      .toolbar { display: none !important; }
      .wrapper { padding: 0 !important; gap: 0 !important; display: block !important; max-width: unset !important; width: auto !important; }
      .page { box-shadow: none; border-radius: 0; margin: 0 auto !important; display: block !important; page-break-after: always !important; break-after: page !important; break-before: avoid-page !important; page-break-inside: avoid !important; }
      .page:last-child { page-break-after: auto !important; break-after: auto !important; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <button id="printBtn">Print</button>
  </div>
  <div class="wrapper" id="pages"></div>
  <script src="https://code.highcharts.com/highcharts.js"><\/script>
</body>
</html>`;
        const win = window.open('', '_blank');
        if (!win) {
          showStatus('Popup blocked. Allow pop-ups to view the booklet.', 'warning');
          return;
        }
        if (typeof computeAxisBounds === 'function') {
          try {
            win.computeAxisBounds = computeAxisBounds;
          } catch (error) {
            /* ignore cross-window assignment issues */
          }
        }
        win.document.open();
        win.document.write(bookletHtml);
        win.document.close();
        win.__BOOKLET_DATA = payload;
        const inject = () => {
          const runtimeEl = win.document.createElement('script');
          runtimeEl.type = 'text/javascript';
          runtimeEl.textContent = '(' + sharedChildRuntimeSource.toString() + ')();';
          win.document.body.appendChild(runtimeEl);

          const scriptEl = win.document.createElement('script');
          scriptEl.type = 'text/javascript';
          scriptEl.textContent = '(' + childRenderSource.toString() + ')();';
          win.document.body.appendChild(scriptEl);
        };
        if (win.document.readyState === 'complete') {
          inject();
        } else {
          win.addEventListener('load', inject, { once: true });
        }
      }

      function openGraphBook(payload) {
        const safeTitle = escapeHtml(payload.title || 'Weekly Graphs');
    const pageInfo = resolvePageSizeInfo(payload.pageSize);
    const pageWidth = pageInfo.width;
    const pageHeight = pageInfo.height;
    const pageSizeCss = `${pageWidth} ${pageHeight}`;
  const pagePadding = resolvePagePadding(pageInfo.key);
  const graphPadding = {
    top: Math.max(pagePadding.top - 0.2, 0.35),
    side: Math.max(pagePadding.side - 0.3, 0.45),
    bottom: Math.max(pagePadding.bottom - 0.15, 0.35)
  };
  const pagePaddingCss = formatPadding(graphPadding);
        const graphHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    :root { --page-width: ${pageWidth}; --page-height: ${pageHeight}; }
  @page { size: ${pageSizeCss}; margin: 0; }
  *, *::before, *::after { box-sizing: border-box; }
  body { margin: 0; background: #ffffff; font-family: 'Segoe UI', Roboto, Arial, sans-serif; display: flex; flex-direction: column; align-items: center; min-height: 100vh; }
  .toolbar { position: sticky; top: 0; z-index: 10; display: flex; justify-content: flex-end; gap: 12px; padding: 12px 16px; background: rgba(255,255,255,0.95); border-bottom: 1px solid rgba(15,23,42,0.08); box-shadow: 0 6px 18px -12px rgba(15,23,42,0.2); }
    .toolbar button { border: 1px solid rgba(15,23,42,0.2); border-radius: 999px; padding: 6px 16px; background: #0f172a; color: #fff; font-weight: 600; cursor: pointer; }
  .wrapper { width: 100%; max-width: calc(var(--page-width) + 2in); margin: 0 auto; padding: 36px 24px; display: flex; flex-direction: column; gap: 24px; align-items: center; }
  .page { width: var(--page-width); height: var(--page-height); background: #fff; border-radius: 16px; box-shadow: 0 28px 60px -32px rgba(15,23,42,0.55); padding: ${pagePaddingCss}; display: flex; align-items: stretch; position: relative; overflow: hidden; break-after: page; }
  .page:last-child { break-after: auto; }
  .page-graph-inner { flex: 1; display: flex; flex-direction: column; gap: 20px; border: 1px solid rgba(15,23,42,0.12); border-radius: 14px; padding: 50px 24px 20px; background: #ffffff; height: 100%; box-sizing: border-box; }
  .page-graph-inner.split { flex-direction: row; align-items: stretch; gap: 20px; justify-content: space-between; }
  .page-graph-info { width: 2.8in; display: flex; flex-direction: column; gap: 12px; justify-content: flex-start; align-items: center; text-align: center; color: #0f172a; font-weight: 600; }
  .page-graph-info > * { margin: 0; }
  .page-graph-info .page-graph-overline,
  .page-graph-info .page-graph-title,
  .page-graph-info .page-graph-subtitle,
  .page-graph-info .page-graph-meta { width: 100%; }
  .page-graph-divider { width: 2px; align-self: stretch; background: rgba(15,23,42,0.16); border-radius: 999px; }
  .page-graph-header { display: flex; align-items: flex-start; justify-content: flex-start; gap: 18px; }
  .page-graph-header.simple { border-bottom: 1px solid rgba(15,23,42,0.15); padding-bottom: 18px; }
  .page-graph-header.simple.with-logo { align-items: center; }
  .page-graph-titles { display: flex; flex-direction: column; gap: 6px; }
    .page-graph-overline { font-size: 0.75rem; letter-spacing: 0.32em; text-transform: uppercase; color: rgba(15,23,42,0.55); font-weight: 700; }
  .page-graph-title { margin: 0; font-size: 1rem; letter-spacing: 0.08em; font-weight: 700; color: #0f172a; }
  .page-graph-subtitle { font-size: 0.9rem; letter-spacing: 0.04em; color: rgba(15,23,42,0.7); font-weight: 600; }
  .page-graph-meta { font-size: 0.8rem; letter-spacing: 0.18em; color: rgba(15,23,42,0.55); font-weight: 600; }
  .page-graph-logo { display: flex; align-items: center; justify-content: center; }
  .page-graph-logo img { max-height: 72px; max-width: 100%; object-fit: contain; }
  .page-graph-logo.top { width: 1.75in; max-width: 30%; }
  .page-graph-logo.side { align-self: center; margin-bottom: 0.25in; }
  .page-graph-chart { flex: 1; min-height: 0; display: flex; align-items: stretch; }
    .page-graph-chart .chart { flex: 1; }
  .page-graph-side .page-graph-info { width: 2.45in; align-items: flex-start; text-align: left; }
  .page-graph-side .page-graph-divider { margin-left: -6px; margin-right: 18px; }
  .page-graph-side .page-graph-title,
  .page-graph-side .page-graph-subtitle { font-size: 18px; letter-spacing: 0.05em; }
  .page-graph-side .page-graph-subtitle { font-weight: 600; color: rgba(15,23,42,0.75); }
    @media print {
      body { background: #fff; margin: 0 !important; min-height: auto !important; display: block !important; }
      .toolbar { display: none !important; }
      .wrapper { padding: 0 !important; gap: 0 !important; display: block !important; max-width: unset !important; width: auto !important; }
      .page { box-shadow: none; border-radius: 0; margin: 0 auto !important; display: block !important; page-break-after: always !important; break-after: page !important; break-before: avoid-page !important; page-break-inside: avoid !important; }
      .page:last-child { page-break-after: auto !important; break-after: auto !important; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <button id="printBtn">Print</button>
  </div>
  <div class="wrapper" id="pages"></div>
  <script src="https://code.highcharts.com/highcharts.js"><\/script>
</body>
</html>`;
        const win = window.open('', '_blank');
        if (!win) {
          showStatus('Popup blocked. Allow pop-ups to view the graphs.', 'warning');
          return;
        }
        if (typeof computeAxisBounds === 'function') {
          try {
            win.computeAxisBounds = computeAxisBounds;
          } catch (error) {
            /* ignore cross-window assignment issues */
          }
        }
        win.document.open();
        win.document.write(graphHtml);
        win.document.close();
        win.__BOOKLET_DATA = payload;
        const inject = () => {
          const runtimeEl = win.document.createElement('script');
          runtimeEl.type = 'text/javascript';
          runtimeEl.textContent = '(' + sharedChildRuntimeSource.toString() + ')();';
          win.document.body.appendChild(runtimeEl);

          const scriptEl = win.document.createElement('script');
          scriptEl.type = 'text/javascript';
          scriptEl.textContent = '(' + childGraphRenderSource.toString() + ')();';
          win.document.body.appendChild(scriptEl);
        };
        if (win.document.readyState === 'complete') {
          inject();
        } else {
          win.addEventListener('load', inject, { once: true });
        }
      }

      function handleGenerate() {
        try {
          if (model && model.graphOnly) {
            showStatus('This CSV only supports Graph Mode (no Division column detected). Use “Open Graph Binder” instead.', 'info');
            return;
          }
          const payload = buildBookletPayload();
          openBooklet(payload);
        } catch (err) {
          showStatus(err.message, 'warning');
        }
      }

      function handleGenerateGraphs() {
        try {
          const payload = buildGraphPayload();
          if (payload.logoMode === 'show' && !payload.logo) {
            payload.logoMode = 'none';
            showStatus('Graph mode logo option selected but no logo uploaded; continuing without logo.', 'info');
          }
          openGraphBook(payload);
        } catch (err) {
          showStatus(err.message, 'warning');
        }
      }

      csvInput.addEventListener('change', handleCsvChange);
      logoInput.addEventListener('change', handleLogoChange);
      divisionContainer.addEventListener('click', handleDivisionClick);
      divisionContainer.addEventListener('change', handleDivisionChange);
  divisionContainer.addEventListener('dragstart', handleDragStart);
  divisionContainer.addEventListener('dragover', handleDragOver);
  divisionContainer.addEventListener('drop', handleDrop);
  divisionContainer.addEventListener('dragend', handleDragEnd);
      generateBtn.addEventListener('click', handleGenerate);
      if (downloadSettingsBtn) {
        downloadSettingsBtn.addEventListener('click', downloadSettings);
      }
      if (editCsvBtn) {
        editCsvBtn.addEventListener('click', openCsvEditor);
      }
      if (csvEditorApplyBtn) {
        csvEditorApplyBtn.addEventListener('click', applyCsvEditorChanges);
      }
      if (csvEditorModalEl) {
        csvEditorModalEl.addEventListener('input', handleCsvEditorInput);
        csvEditorModalEl.addEventListener('shown.bs.modal', () => {
          if (!csvEditorTableBody) return;
          const firstInput = csvEditorTableBody.querySelector('.csv-editor-input');
          if (firstInput) {
            firstInput.focus({ preventScroll: false });
          }
        });
        csvEditorModalEl.addEventListener('hidden.bs.modal', () => {
          csvEditorWorking = null;
          csvEditorDirty = false;
          updateCsvEditorApplyState();
          clearCsvEditorAlert();
          if (csvEditorSearchDebounce) {
            clearTimeout(csvEditorSearchDebounce);
            csvEditorSearchDebounce = null;
          }
        });
      }
      if (csvEditorShowGaps) {
        csvEditorShowGaps.addEventListener('change', () => {
          csvEditorFilters.showGaps = csvEditorShowGaps.checked;
          renderCsvEditorTable();
        });
      }
      if (csvEditorSearchInput) {
        csvEditorSearchInput.addEventListener('input', () => {
          csvEditorFilters.search = csvEditorSearchInput.value || '';
          if (csvEditorSearchDebounce) {
            clearTimeout(csvEditorSearchDebounce);
          }
          csvEditorSearchDebounce = setTimeout(() => {
            renderCsvEditorTable();
          }, 180);
        });
      }
      if (pageSizeSelect) {
        pageSizeSelect.addEventListener('change', (event) => {
          pageSize = event.target.value;
        });
      }
      if (defaultThemeSelect) {
        defaultThemeSelect.addEventListener('change', (event) => {
          const value = event.target.value;
          if (COLOR_KEYS.includes(value)) {
            defaultThemeKey = value;
            if (model && model.settings) {
              model.settings.defaultThemeKey = value;
            }
          }
          refreshDefaultThemeLabel();
        });
      }
      if (statTitleModeSelect) {
        statTitleModeSelect.addEventListener('change', (event) => {
          statTitleMode = event.target.value;
          renderDivisions(true);
        });
      }
      if (graphHeadingSelect) {
        graphHeadingSelect.addEventListener('change', (event) => {
          graphHeadingMode = event.target.value;
        });
      }
      if (graphLogoSelect) {
        graphLogoSelect.addEventListener('change', (event) => {
          graphLogoMode = event.target.value;
        });
      }
      if (graphLayoutSelect) {
        graphLayoutSelect.addEventListener('change', (event) => {
          graphLayoutMode = event.target.value;
        });
      }
      updateCsvEditorApplyState();
      if (generateGraphsBtn) {
        generateGraphsBtn.addEventListener('click', handleGenerateGraphs);
      }
      if (divisionSettingsModalEl) {
        divisionSettingsModalEl.addEventListener('shown.bs.modal', () => {
          if (divisionModalFocusField === 'color' && divisionColorSelect) {
            divisionColorSelect.focus();
            divisionModalFocusField = null;
          }
        });
        divisionSettingsModalEl.addEventListener('hidden.bs.modal', () => {
          editingDivisionIndex = null;
          divisionModalFocusField = null;
        });
      }
      if (divisionSettingsSave) {
        divisionSettingsSave.addEventListener('click', () => {
          if (!model || editingDivisionIndex == null) {
            if (divisionModal) divisionModal.hide();
            return;
          }
          const division = model.divisions[editingDivisionIndex];
          if (!division) {
            if (divisionModal) divisionModal.hide();
            return;
          }
          if (divisionVfpInput) {
            division.vfpText = divisionVfpInput.value;
          }
          if (divisionColorSelect) {
            const key = divisionColorSelect.value;
            if (key && COLOR_THEMES[key]) {
              const theme = COLOR_THEMES[key];
              division.color = { ...theme };
              division.colorKey = theme.key;
            }
          }
          let pendingAccent = null;
          let resetAccent = false;
          let initialAccent = null;
          if (divisionCustomColorInput) {
            if (divisionCustomColorInput.dataset.reset === 'true') {
              resetAccent = true;
              delete divisionCustomColorInput.dataset.reset;
            } else {
              pendingAccent = normalizeHexColor(divisionCustomColorInput.value);
            }
            if (divisionCustomColorInput.dataset.initial) {
              initialAccent = normalizeHexColor(divisionCustomColorInput.dataset.initial);
            }
          }
          if (resetAccent) {
            division.customAccent = null;
            const theme = COLOR_THEMES[division.colorKey] || COLOR_THEMES[COLOR_KEYS[0]];
            if (theme) {
              division.color = { ...theme };
            }
          } else if (pendingAccent) {
            const palette = paletteFromAccent(pendingAccent);
            const accentChanged = pendingAccent !== initialAccent;
            if (palette && (division.customAccent || accentChanged)) {
              division.customAccent = palette.accent;
              division.color = { ...division.color, ...palette, key: division.colorKey };
            }
          } else if (!division.customAccent) {
            division.customAccent = null;
          }
          if (divisionCustomColorReset) {
            divisionCustomColorReset.disabled = !division.customAccent;
          }
          if (divisionVfpAlignSelect) {
            const align = divisionVfpAlignSelect.value;
            division.vfpAlign = ['left', 'right', 'center'].includes(align) ? align : 'center';
          }
          if (divisionVfpWidthSelect) {
            const width = divisionVfpWidthSelect.value;
            division.vfpWidth = ['standard', 'wide', 'narrow'].includes(width) ? width : 'wide';
          }
          if (divisionVfpScaleSelect) {
            const scale = divisionVfpScaleSelect.value;
            division.vfpScale = scale === 'compact' || scale === 'large' || scale === 'normal' ? scale : 'normal';
          }
          if (divisionVfpFontSelect) {
            const font = divisionVfpFontSelect.value;
            division.vfpFont = ['small', 'medium', 'large', 'xlarge'].includes(font) ? font : 'large';
          }
          if (divisionCustomColorDefault && divisionCustomColorDefault.checked) {
            divisionCustomColorDefault.checked = false;
            if (division.customAccent) {
              defaultCustomAccent = division.customAccent;
              if (model.settings) {
                model.settings.defaultCustomAccent = defaultCustomAccent;
              }
              const palette = paletteFromAccent(defaultCustomAccent);
              if (palette) {
                model.divisions.forEach((otherDivision, idx) => {
                  if (idx === editingDivisionIndex) return;
                  if (otherDivision.customAccent) return;
                  otherDivision.color = { ...otherDivision.color, ...palette, key: otherDivision.colorKey };
                });
              }
            }
          }
          if (divisionModal) divisionModal.hide();
          renderDivisions(true);
        });
      }
      if (divisionColorSelect && divisionCustomColorInput) {
        divisionColorSelect.addEventListener('change', () => {
          if (divisionCustomColorReset && !divisionCustomColorReset.disabled) {
            return;
          }
          const theme = COLOR_THEMES[divisionColorSelect.value];
          if (theme) {
            const accentHex = normalizeHexColor(theme.accent) || '#1D4ED8';
            divisionCustomColorInput.value = accentHex;
            divisionCustomColorInput.dataset.initial = accentHex;
            if (divisionCustomColorReset) {
              divisionCustomColorReset.disabled = true;
            }
          }
        });
      }
      if (divisionCustomColorReset && divisionCustomColorInput) {
        divisionCustomColorReset.addEventListener('click', () => {
          divisionCustomColorInput.dataset.reset = 'true';
          if (divisionColorSelect) {
            const key = divisionColorSelect.value;
            const theme = COLOR_THEMES[key] || COLOR_THEMES[COLOR_KEYS[0]];
            if (theme) {
              divisionCustomColorInput.value = normalizeHexColor(theme.accent) || '#1D4ED8';
            }
          }
          if (divisionCustomColorDefault) {
            divisionCustomColorDefault.checked = false;
          }
          divisionCustomColorReset.disabled = true;
        });
      }
      if (divisionCustomColorInput) {
        divisionCustomColorInput.addEventListener('input', () => {
          if (divisionCustomColorInput.dataset.reset) {
            delete divisionCustomColorInput.dataset.reset;
          }
          if (divisionCustomColorReset) {
            divisionCustomColorReset.disabled = false;
          }
        });
      }
  });
