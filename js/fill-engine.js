(function (global) {
  'use strict'

  const PREFIX = 'WPS_AutoFill_'
  const PREVIEW_PREFIX = PREFIX + 'Preview_'
  const PICKED_PREVIEW_PREFIX = PREVIEW_PREFIX + 'Picked_'
  let lastPreviewSelectionKey = ''
  const M = {
    msoFalse: 0,
    msoTrue: -1,
    msoAutoShape: 1,
    msoFreeform: 5,
    msoGroup: 6,
    msoEditingCorner: 1,
    msoSegmentLine: 0,
    msoSegmentCurve: 1,
    msoSendToBack: 1,
    msoShapeRectangle: 1,
    msoShapeParallelogram: 2,
    msoShapeTrapezoid: 3,
    msoShapeDiamond: 4,
    msoShapeRoundedRectangle: 5,
    msoShapeOctagon: 6,
    msoShapeIsoscelesTriangle: 7,
    msoShapeRightTriangle: 8,
    msoShapeOval: 9,
    msoShapeHexagon: 10,
    msoShapeTypeLine: 9,
    msoConnectorStraight: 1
  }

  function getApplication() {
    if (global.Application && global.Application.ActivePresentation !== undefined) {
      return global.Application
    }
    if (global.wps && typeof global.wps.WppApplication === 'function') {
      return global.wps.WppApplication()
    }
    throw new Error('未检测到 WPS 演示加载项环境。')
  }

  function toNumber(value, fallback) {
    const n = Number(value)
    return Number.isFinite(n) ? n : fallback
  }

  function isTrue(value) {
    return value === true || Number(value) === M.msoTrue || Number(value) === 1
  }

  function shapeName(shape) {
    try { return String(shape.Name || '') } catch (_) { return '' }
  }

  function isGeneratedShape(shape) {
    const name = shapeName(shape)
    return name.indexOf(PREFIX) === 0 && name.indexOf(PREVIEW_PREFIX) !== 0
  }

  function isPreviewShape(shape) {
    return shapeName(shape).indexOf(PREVIEW_PREFIX) === 0
  }

  function isPickedPreviewShape(shape) {
    return shapeName(shape).indexOf(PICKED_PREVIEW_PREFIX) === 0
  }

  function isPluginShape(shape) {
    return shapeName(shape).indexOf(PREFIX) === 0
  }

  function getAutoShapeType(shape) {
    try { return Number(shape.AutoShapeType) } catch (_) { return NaN }
  }

  function isEllipse(shape) {
    if (getAutoShapeType(shape) === M.msoShapeOval) return true
    return /^(椭圆|oval|ellipse)/i.test(shapeName(shape))
  }

  function isStraightLine(shape) {
    let type = NaN
    try { type = Number(shape.Type) } catch (_) { return false }
    if (type !== M.msoShapeTypeLine) return false

    try {
      if (isTrue(shape.Connector)) {
        const connectorType = Number(shape.ConnectorFormat.Type)
        return connectorType === M.msoConnectorStraight
      }
    } catch (_) {
      // 普通直线通常没有可用的 ConnectorFormat。
    }
    return true
  }

  function getSelectedShapes(app, emptyMessage, allowEmpty) {
    const selection = app.ActiveWindow && app.ActiveWindow.Selection
    if (!selection) {
      if (allowEmpty) return []
      throw new Error('当前窗口没有可用的选择对象。')
    }

    let range
    try { range = selection.ShapeRange } catch (_) {
      range = null
    }
    if (!range) {
      if (allowEmpty) return []
      throw new Error(emptyMessage || '当前没有选中任何形状，请先在幻灯片上选中对象。')
    }

    const result = []
    let count = 0
    try { count = toNumber(range.Count, 0) } catch (_) {}
    if (count < 1) {
      if (allowEmpty) return []
      throw new Error(emptyMessage || '当前没有选中任何形状，请先在幻灯片上选中对象。')
    }
    for (let i = 1; i <= count; i += 1) result.push(range.Item(i))
    return result
  }

  function expandShape(shape, result, depth) {
    if (!shape || depth > 8) return
    let type = NaN
    try { type = Number(shape.Type) } catch (_) {}
    if (type === M.msoGroup) {
      let items = null
      try { items = shape.GroupItems } catch (_) {}
      const count = toNumber(items && items.Count, 0)
      if (count > 0) {
        for (let i = 1; i <= count; i += 1) {
          try { expandShape(items.Item(i), result, depth + 1) } catch (_) {}
        }
        return
      }
    }
    result.push(shape)
  }

  function expandShapes(shapes) {
    const result = []
    ;(shapes || []).forEach(shape => expandShape(shape, result, 0))
    return result
  }

  function getSlideShapes(slide) {
    const result = []
    let count = 0
    try { count = toNumber(slide && slide.Shapes && slide.Shapes.Count, 0) } catch (_) {}
    for (let i = 1; i <= count; i += 1) {
      const shape = slide.Shapes.Item(i)
      const expanded = expandShapes([shape])
      expanded.forEach(item => {
        if (!isPluginShape(item)) result.push(item)
      })
    }
    return result
  }

  function getActiveSlide(app) {
    try {
      const slide = app.ActiveWindow.View.Slide
      if (slide) return slide
    } catch (_) {}

    try {
      const index = app.ActiveWindow.View.Slide.SlideIndex
      return app.ActivePresentation.Slides.Item(index)
    } catch (_) {}

    throw new Error('无法取得当前幻灯片，请切换到普通编辑视图。')
  }

  function unselectCurrentShapes(app) {
    try {
      const selection = app.ActiveWindow && app.ActiveWindow.Selection
      if (selection && typeof selection.Unselect === 'function') selection.Unselect()
    } catch (_) {}
  }

  function ellipseFromShape(shape) {
    const left = toNumber(shape.Left, 0)
    const top = toNumber(shape.Top, 0)
    const width = toNumber(shape.Width, 0)
    const height = toNumber(shape.Height, 0)
    if (width <= 0 || height <= 0) throw new Error('椭圆尺寸无效。')
    return {
      shape,
      cx: left + width / 2,
      cy: top + height / 2,
      rx: width / 2,
      ry: height / 2,
      left,
      top,
      width,
      height
    }
  }

  function getShapeType(shape) {
    try { return Number(shape.Type) } catch (_) { return NaN }
  }

  function shapeBounds(shape) {
    const left = toNumber(shape.Left, 0)
    const top = toNumber(shape.Top, 0)
    const width = Math.abs(toNumber(shape.Width, 0))
    const height = Math.abs(toNumber(shape.Height, 0))
    if (width <= 1e-8 || height <= 1e-8) {
      throw new Error('闭合外框尺寸无效：' + shapeName(shape))
    }
    return {
      left,
      top,
      width,
      height,
      cx: left + width / 2,
      cy: top + height / 2
    }
  }

  function shapeTransform(shape, bounds) {
    let hFlip = false
    let vFlip = false
    try { hFlip = isTrue(shape.HorizontalFlip) } catch (_) {}
    try { vFlip = isTrue(shape.VerticalFlip) } catch (_) {}
    const angle = toNumber(shape.Rotation, 0) * Math.PI / 180
    const cos = Math.cos(angle)
    const sin = Math.sin(angle)

    return function transform(nx, ny) {
      let x = bounds.left + (hFlip ? 1 - nx : nx) * bounds.width
      let y = bounds.top + (vFlip ? 1 - ny : ny) * bounds.height
      if (Math.abs(angle) > 1e-10) {
        const dx = x - bounds.cx
        const dy = y - bounds.cy
        x = bounds.cx + dx * cos - dy * sin
        y = bounds.cy + dx * sin + dy * cos
      }
      return { x, y }
    }
  }

  function sampleNormalizedEllipse(transform, segments) {
    const points = []
    for (let i = 0; i < segments; i += 1) {
      const angle = Math.PI * 2 * i / segments
      points.push(transform(0.5 + 0.5 * Math.cos(angle), 0.5 + 0.5 * Math.sin(angle)))
    }
    return points
  }

  function sampleNormalizedRoundedRectangle(transform, segments, radiusX, radiusY) {
    const points = []
    const rx = Math.max(0.01, Math.min(0.49, radiusX))
    const ry = Math.max(0.01, Math.min(0.49, radiusY))
    const arcSteps = Math.max(4, Math.round(segments / 4))
    const addArc = (cx, cy, startAngle) => {
      for (let i = 0; i <= arcSteps; i += 1) {
        const angle = startAngle + Math.PI / 2 * i / arcSteps
        pushDistinct(points, transform(
          cx + rx * Math.cos(angle),
          cy + ry * Math.sin(angle)
        ), 1e-8)
      }
    }
    addArc(1 - rx, ry, -Math.PI / 2)
    addArc(1 - rx, 1 - ry, 0)
    addArc(rx, 1 - ry, Math.PI / 2)
    addArc(rx, ry, Math.PI)
    return points
  }

  function normalizedPolygon(autoShapeType) {
    switch (autoShapeType) {
      case M.msoShapeRectangle:
        return [[0, 0], [1, 0], [1, 1], [0, 1]]
      case M.msoShapeParallelogram:
        return [[0.22, 0], [1, 0], [0.78, 1], [0, 1]]
      case M.msoShapeTrapezoid:
        return [[0.22, 0], [0.78, 0], [1, 1], [0, 1]]
      case M.msoShapeDiamond:
        return [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]
      case M.msoShapeOctagon:
        return [[0.29, 0], [0.71, 0], [1, 0.29], [1, 0.71], [0.71, 1], [0.29, 1], [0, 0.71], [0, 0.29]]
      case M.msoShapeIsoscelesTriangle:
        return [[0.5, 0], [1, 1], [0, 1]]
      case M.msoShapeRightTriangle:
        return [[0, 0], [1, 1], [0, 1]]
      case M.msoShapeHexagon:
        return [[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]]
      default:
        return null
    }
  }

  function flattenNumbers(value, output, depth) {
    if (depth > 6 || value === null || value === undefined) return
    if (typeof value === 'number' && Number.isFinite(value)) {
      output.push(value)
      return
    }
    if (Array.isArray(value)) {
      value.forEach(item => flattenNumbers(item, output, depth + 1))
      return
    }
    if (typeof value === 'object' && Number.isFinite(Number(value.length))) {
      for (let i = 0; i < Number(value.length); i += 1) {
        flattenNumbers(value[i], output, depth + 1)
      }
      return
    }
    if (typeof value === 'object' && Number.isFinite(Number(value.Count)) && typeof value.Item === 'function') {
      for (let i = 1; i <= Number(value.Count); i += 1) {
        try { flattenNumbers(value.Item(i), output, depth + 1) } catch (_) {}
      }
      return
    }
    if (typeof value === 'object') {
      Object.keys(value)
        .filter(key => /^\d+$/.test(key))
        .sort((a, b) => Number(a) - Number(b))
        .forEach(key => flattenNumbers(value[key], output, depth + 1))
    }
  }

  function pointFromShapeNode(node) {
    let raw
    try { raw = node.Points } catch (_) { return null }
    const numbers = []
    flattenNumbers(raw, numbers, 0)
    if (numbers.length >= 2) {
      return { x: numbers[numbers.length - 2], y: numbers[numbers.length - 1] }
    }
    return null
  }

  function pointPairsFromValue(value) {
    const numbers = []
    flattenNumbers(value, numbers, 0)
    const points = []
    for (let i = 0; i + 1 < numbers.length; i += 2) {
      points.push({ x: numbers[i], y: numbers[i + 1] })
    }
    return points
  }

  function cubicBezierPoint(p0, p1, p2, p3, t) {
    const mt = 1 - t
    const mt2 = mt * mt
    const t2 = t * t
    return {
      x: p0.x * mt2 * mt + 3 * p1.x * mt2 * t + 3 * p2.x * mt * t2 + p3.x * t2 * t,
      y: p0.y * mt2 * mt + 3 * p1.y * mt2 * t + 3 * p2.y * mt * t2 + p3.y * t2 * t
    }
  }

  function sampleFreeformPath(shape, segments) {
    let nodes = null
    try { nodes = shape.Nodes } catch (_) {}
    const nodeCount = toNumber(nodes && nodes.Count, 0)
    let hasCurve = false
    const nodePoints = []
    for (let i = 1; i <= nodeCount; i += 1) {
      const node = nodes.Item(i)
      try {
        if (Number(node.SegmentType) === M.msoSegmentCurve) hasCurve = true
      } catch (_) {}
      const point = pointFromShapeNode(node)
      if (point) pushDistinct(nodePoints, point, 1e-6)
    }

    let vertices = []
    try { vertices = pointPairsFromValue(shape.Vertices) } catch (_) {}
    if (vertices.length < 2) vertices = nodePoints
    if (vertices.length < 2) {
      throw new Error('无法读取自由线条节点：' + shapeName(shape))
    }

    const xs = vertices.map(point => point.x)
    const ys = vertices.map(point => point.y)
    const scale = Math.max(
      Math.max.apply(null, xs) - Math.min.apply(null, xs),
      Math.max.apply(null, ys) - Math.min.apply(null, ys),
      1
    )
    const epsilon = scale * 1e-6
    const closed = Math.hypot(
      vertices[0].x - vertices[vertices.length - 1].x,
      vertices[0].y - vertices[vertices.length - 1].y
    ) <= epsilon

    let points = []
    if (hasCurve && vertices.length >= 4 && (vertices.length - 1) % 3 === 0) {
      const curveCount = (vertices.length - 1) / 3
      const steps = Math.max(6, Math.min(48, Math.ceil(toNumber(segments, 160) / Math.max(curveCount, 1))))
      pushDistinct(points, vertices[0], epsilon)
      for (let curve = 0; curve < curveCount; curve += 1) {
        const offset = curve * 3
        for (let step = 1; step <= steps; step += 1) {
          pushDistinct(points, cubicBezierPoint(
            vertices[offset],
            vertices[offset + 1],
            vertices[offset + 2],
            vertices[offset + 3],
            step / steps
          ), epsilon)
        }
      }
    } else {
      points = vertices.slice()
    }

    if (closed && points.length > 1) points.pop()
    points = closed ? cleanPolygon(points, epsilon) : points.reduce((result, point) => {
      pushDistinct(result, point, epsilon)
      return result
    }, [])
    return { shape, name: shapeName(shape), points, closed, curved: hasCurve }
  }

  function sampleFreeform(shape) {
    const path = sampleFreeformPath(shape, 160)
    if (!path.closed || path.points.length < 3) {
      throw new Error('自由形状必须是已闭合、至少包含 3 个节点的轮廓：' + shapeName(shape))
    }
    return path.points
  }

  function supportedBoundaryKind(shape) {
    if (isPluginShape(shape) || isStraightLine(shape)) return ''
    const type = getShapeType(shape)
    const autoShapeType = getAutoShapeType(shape)
    if (autoShapeType === M.msoShapeOval) return '椭圆'
    if (autoShapeType === M.msoShapeRoundedRectangle) return '圆角矩形'
    if (normalizedPolygon(autoShapeType)) return '多边形'
    if (type === M.msoFreeform) return '自由形状'
    return ''
  }

  function boundaryFromShape(shape, segments) {
    const kind = supportedBoundaryKind(shape)
    if (!kind) {
      throw new Error('不支持的外框类型：' + shapeName(shape))
    }

    let points
    if (kind === '自由形状') {
      points = sampleFreeform(shape)
    } else {
      const bounds = shapeBounds(shape)
      const transform = shapeTransform(shape, bounds)
      if (kind === '椭圆') {
        points = sampleNormalizedEllipse(transform, segments)
      } else if (kind === '圆角矩形') {
        const radius = Math.min(bounds.width, bounds.height) * 0.18
        points = sampleNormalizedRoundedRectangle(
          transform,
          Math.max(24, Math.round(segments / 2)),
          radius / bounds.width,
          radius / bounds.height
        )
      } else {
        points = normalizedPolygon(getAutoShapeType(shape)).map(point => transform(point[0], point[1]))
      }
    }

    const area = Math.abs(polygonArea(points))
    if (area <= 1e-6) throw new Error('闭合外框面积过小：' + shapeName(shape))
    return { shape, name: shapeName(shape), kind, points, area }
  }

  function lineFromShape(shape) {
    const rotation = ((toNumber(shape.Rotation, 0) % 360) + 360) % 360
    const left = toNumber(shape.Left, 0)
    const top = toNumber(shape.Top, 0)
    const width = Math.abs(toNumber(shape.Width, 0))
    const height = Math.abs(toNumber(shape.Height, 0))
    if (width < 1e-8 && height < 1e-8) throw new Error('检测到零长度直线：' + shapeName(shape))

    let hFlip = false
    let vFlip = false
    try { hFlip = isTrue(shape.HorizontalFlip) } catch (_) {}
    try { vFlip = isTrue(shape.VerticalFlip) } catch (_) {}

    const negativeSlope = hFlip !== vFlip
    let p1 = negativeSlope
      ? { x: left + width, y: top }
      : { x: left, y: top }
    let p2 = negativeSlope
      ? { x: left, y: top + height }
      : { x: left + width, y: top + height }

    if (Math.min(rotation, Math.abs(rotation - 360)) > 0.05) {
      const cx = left + width / 2
      const cy = top + height / 2
      const angle = rotation * Math.PI / 180
      const cos = Math.cos(angle)
      const sin = Math.sin(angle)
      const rotate = point => {
        const dx = point.x - cx
        const dy = point.y - cy
        return {
          x: cx + dx * cos - dy * sin,
          y: cy + dx * sin + dy * cos
        }
      }
      p1 = rotate(p1)
      p2 = rotate(p2)
    }

    return { shape, p1, p2, name: shapeName(shape) }
  }

  function segmentEllipseRoots(line, ellipse) {
    const dx = line.p2.x - line.p1.x
    const dy = line.p2.y - line.p1.y
    const x = line.p1.x - ellipse.cx
    const y = line.p1.y - ellipse.cy
    const rx2 = ellipse.rx * ellipse.rx
    const ry2 = ellipse.ry * ellipse.ry

    const a = (dx * dx) / rx2 + (dy * dy) / ry2
    const b = 2 * ((x * dx) / rx2 + (y * dy) / ry2)
    const c = (x * x) / rx2 + (y * y) / ry2 - 1
    const disc = b * b - 4 * a * c
    if (a <= 1e-15 || disc <= 1e-10) return []

    const root = Math.sqrt(disc)
    return [(-b - root) / (2 * a), (-b + root) / (2 * a)].sort((m, n) => m - n)
  }

  function validateDivider(line, ellipse) {
    const roots = segmentEllipseRoots(line, ellipse)
    if (roots.length !== 2) {
      throw new Error('直线“' + line.name + '”没有穿过椭圆内部。')
    }
    const tolerance = 0.03
    if (roots[0] < -tolerance || roots[1] > 1 + tolerance) {
      throw new Error('直线“' + line.name + '”没有完整贯穿椭圆；请让两端稍微伸出椭圆边界。')
    }
    return roots
  }

  function sampleEllipse(ellipse, segments) {
    const points = []
    for (let i = 0; i < segments; i += 1) {
      const angle = (Math.PI * 2 * i) / segments
      points.push({
        x: ellipse.cx + ellipse.rx * Math.cos(angle),
        y: ellipse.cy + ellipse.ry * Math.sin(angle)
      })
    }
    return points
  }

  function crossLine(line, point) {
    const dx = line.p2.x - line.p1.x
    const dy = line.p2.y - line.p1.y
    return dx * (point.y - line.p1.y) - dy * (point.x - line.p1.x)
  }

  function intersectionOnLine(a, b, da, db) {
    const denom = da - db
    const t = Math.abs(denom) < 1e-15 ? 0.5 : da / denom
    return {
      x: a.x + (b.x - a.x) * t,
      y: a.y + (b.y - a.y) * t
    }
  }

  function pushDistinct(points, point, epsilon) {
    if (!points.length) {
      points.push(point)
      return
    }
    const last = points[points.length - 1]
    if (Math.hypot(last.x - point.x, last.y - point.y) > epsilon) points.push(point)
  }

  function cleanPolygon(points, epsilon) {
    if (!points || points.length < 3) return []
    const cleaned = []
    for (const p of points) pushDistinct(cleaned, p, epsilon)
    if (cleaned.length > 1) {
      const first = cleaned[0]
      const last = cleaned[cleaned.length - 1]
      if (Math.hypot(first.x - last.x, first.y - last.y) <= epsilon) cleaned.pop()
    }
    if (cleaned.length < 3) return []

    const reduced = []
    for (let i = 0; i < cleaned.length; i += 1) {
      const prev = cleaned[(i - 1 + cleaned.length) % cleaned.length]
      const curr = cleaned[i]
      const next = cleaned[(i + 1) % cleaned.length]
      const area2 = Math.abs((curr.x - prev.x) * (next.y - curr.y) - (curr.y - prev.y) * (next.x - curr.x))
      const scale = Math.hypot(curr.x - prev.x, curr.y - prev.y) + Math.hypot(next.x - curr.x, next.y - curr.y)
      if (area2 > epsilon * Math.max(1, scale)) reduced.push(curr)
    }
    return reduced.length >= 3 ? reduced : cleaned
  }

  function clipHalfPlane(poly, line, keepPositive, epsilon) {
    const out = []
    for (let i = 0; i < poly.length; i += 1) {
      const a = poly[i]
      const b = poly[(i + 1) % poly.length]
      const da = crossLine(line, a)
      const db = crossLine(line, b)
      const insideA = keepPositive ? da >= -epsilon : da <= epsilon
      const insideB = keepPositive ? db >= -epsilon : db <= epsilon

      if (insideA) pushDistinct(out, a, epsilon)
      if (insideA !== insideB) pushDistinct(out, intersectionOnLine(a, b, da, db), epsilon)
    }
    return cleanPolygon(out, epsilon)
  }

  function polygonArea(poly) {
    let sum = 0
    for (let i = 0; i < poly.length; i += 1) {
      const a = poly[i]
      const b = poly[(i + 1) % poly.length]
      sum += a.x * b.y - b.x * a.y
    }
    return sum / 2
  }

  function polygonCentroid(poly) {
    let x = 0
    let y = 0
    let factorSum = 0
    for (let i = 0; i < poly.length; i += 1) {
      const a = poly[i]
      const b = poly[(i + 1) % poly.length]
      const factor = a.x * b.y - b.x * a.y
      factorSum += factor
      x += (a.x + b.x) * factor
      y += (a.y + b.y) * factor
    }
    if (Math.abs(factorSum) < 1e-12) {
      const average = poly.reduce((acc, p) => ({ x: acc.x + p.x, y: acc.y + p.y }), { x: 0, y: 0 })
      return { x: average.x / poly.length, y: average.y / poly.length }
    }
    return { x: x / (3 * factorSum), y: y / (3 * factorSum) }
  }

  function splitPolygon(poly, line, epsilon, minArea) {
    let hasPositive = false
    let hasNegative = false
    for (const point of poly) {
      const d = crossLine(line, point)
      if (d > epsilon) hasPositive = true
      if (d < -epsilon) hasNegative = true
    }
    if (!hasPositive || !hasNegative) return [poly]

    const positive = clipHalfPlane(poly, line, true, epsilon)
    const negative = clipHalfPlane(poly, line, false, epsilon)
    const pieces = []
    if (positive.length >= 3 && Math.abs(polygonArea(positive)) >= minArea) pieces.push(positive)
    if (negative.length >= 3 && Math.abs(polygonArea(negative)) >= minArea) pieces.push(negative)
    return pieces.length === 2 ? pieces : [poly]
  }

  function crossVector(a, b) {
    return a.x * b.y - a.y * b.x
  }

  function subtractPoints(a, b) {
    return { x: a.x - b.x, y: a.y - b.y }
  }

  function interpolatePoint(a, b, t) {
    return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t }
  }

  function pointOnSegment(point, a, b, epsilon) {
    const ab = subtractPoints(b, a)
    const ap = subtractPoints(point, a)
    if (Math.abs(crossVector(ab, ap)) > epsilon * Math.max(1, Math.hypot(ab.x, ab.y))) return false
    const dot = ap.x * ab.x + ap.y * ab.y
    const length2 = ab.x * ab.x + ab.y * ab.y
    return dot >= -epsilon && dot <= length2 + epsilon
  }

  function pointInPolygon(point, polygon, epsilon) {
    let inside = false
    for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
      const a = polygon[j]
      const b = polygon[i]
      if (pointOnSegment(point, a, b, epsilon)) return true
      const crosses = ((a.y > point.y) !== (b.y > point.y)) &&
        (point.x < (b.x - a.x) * (point.y - a.y) / (b.y - a.y) + a.x)
      if (crosses) inside = !inside
    }
    return inside
  }

  function lineEdgeIntersection(line, a, b, epsilon) {
    const r = subtractPoints(line.p2, line.p1)
    const s = subtractPoints(b, a)
    const qp = subtractPoints(a, line.p1)
    const denominator = crossVector(r, s)
    if (Math.abs(denominator) <= epsilon) {
      return { collinear: Math.abs(crossVector(qp, r)) <= epsilon }
    }
    return {
      collinear: false,
      t: crossVector(qp, s) / denominator,
      u: crossVector(qp, r) / denominator
    }
  }

  function uniqueSortedNumbers(values, epsilon) {
    const sorted = values.slice().sort((a, b) => a - b)
    const result = []
    for (const value of sorted) {
      if (!result.length || Math.abs(value - result[result.length - 1]) > epsilon) result.push(value)
    }
    return result
  }

  function dividerSegmentsInsideBoundary(line, boundary, epsilon) {
    const intersections = []
    let hasPositive = false
    let hasNegative = false
    for (const point of boundary) {
      const side = crossLine(line, point)
      if (side > epsilon) hasPositive = true
      if (side < -epsilon) hasNegative = true
    }
    if (!hasPositive || !hasNegative) {
      throw new Error('直线“' + line.name + '”没有切开所选外框。')
    }

    for (let i = 0; i < boundary.length; i += 1) {
      const a = boundary[i]
      const b = boundary[(i + 1) % boundary.length]
      const hit = lineEdgeIntersection(line, a, b, epsilon)
      if (hit.collinear) {
        throw new Error('直线“' + line.name + '”与外框边缘重合，无法确定填色区域。')
      }
      if (Number.isFinite(hit.t) && hit.u >= -epsilon && hit.u <= 1 + epsilon) {
        intersections.push(hit.t)
      }
    }

    const length = Math.hypot(line.p2.x - line.p1.x, line.p2.y - line.p1.y)
    const parameterEpsilon = Math.max(1e-9, epsilon / Math.max(length, 1))
    const roots = uniqueSortedNumbers(intersections, parameterEpsilon)
    if (roots.length < 2) {
      throw new Error('直线“' + line.name + '”没有完整穿过所选外框。')
    }
    const tolerance = 0.03
    if (roots[0] < -tolerance || roots[roots.length - 1] > 1 + tolerance) {
      throw new Error('直线“' + line.name + '”没有完整贯穿外框；请让两端稍微伸出外框边界。')
    }

    const pieces = []
    for (let i = 0; i < roots.length - 1; i += 1) {
      const start = Math.max(0, roots[i])
      const end = Math.min(1, roots[i + 1])
      if (end - start <= parameterEpsilon) continue
      const middle = interpolatePoint(line.p1, line.p2, (start + end) / 2)
      if (pointInPolygon(middle, boundary, epsilon)) {
        pieces.push({
          a: interpolatePoint(line.p1, line.p2, start),
          b: interpolatePoint(line.p1, line.p2, end),
          kind: 'divider',
          name: line.name
        })
      }
    }
    if (!pieces.length) {
      throw new Error('直线“' + line.name + '”在外框内部没有形成有效分割。')
    }
    return pieces
  }

  function segmentIntersection(segmentA, segmentB, epsilon) {
    const r = subtractPoints(segmentA.b, segmentA.a)
    const s = subtractPoints(segmentB.b, segmentB.a)
    const qp = subtractPoints(segmentB.a, segmentA.a)
    const denominator = crossVector(r, s)
    if (Math.abs(denominator) <= epsilon) {
      return { collinear: Math.abs(crossVector(qp, r)) <= epsilon }
    }
    const t = crossVector(qp, s) / denominator
    const u = crossVector(qp, r) / denominator
    if (t < -epsilon || t > 1 + epsilon || u < -epsilon || u > 1 + epsilon) return null
    return {
      collinear: false,
      t: Math.max(0, Math.min(1, t)),
      u: Math.max(0, Math.min(1, u))
    }
  }

  function genericPathFromShape(shape, segments) {
    if (isPluginShape(shape)) return null
    if (isStraightLine(shape)) {
      const line = lineFromShape(shape)
      return { shape, name: line.name, points: [line.p1, line.p2], closed: false, curved: false }
    }
    if (getShapeType(shape) === M.msoFreeform) {
      return sampleFreeformPath(shape, segments)
    }
    if (supportedBoundaryKind(shape)) {
      const boundary = boundaryFromShape(shape, segments)
      return {
        shape,
        name: boundary.name,
        points: boundary.points,
        closed: true,
        curved: boundary.kind === '椭圆' || boundary.kind === '圆角矩形'
      }
    }
    return null
  }

  function segmentParameter(point, segment) {
    const dx = segment.b.x - segment.a.x
    const dy = segment.b.y - segment.a.y
    if (Math.abs(dx) >= Math.abs(dy) && Math.abs(dx) > 1e-12) return (point.x - segment.a.x) / dx
    if (Math.abs(dy) > 1e-12) return (point.y - segment.a.y) / dy
    return 0
  }

  function addCollinearSegmentSplits(a, b, epsilon) {
    for (const point of [a.a, a.b]) {
      if (pointOnSegment(point, b.a, b.b, epsilon)) b.splits.push(segmentParameter(point, b))
    }
    for (const point of [b.a, b.b]) {
      if (pointOnSegment(point, a.a, a.b, epsilon)) a.splits.push(segmentParameter(point, a))
    }
  }

  function buildGenericRegions(paths) {
    const usablePaths = (paths || []).filter(path => {
      return path && path.points && path.points.length >= (path.closed ? 3 : 2)
    })
    if (!usablePaths.length) throw new Error('当前页没有可读取的线条或闭合形状。')

    const allPoints = usablePaths.reduce((result, path) => result.concat(path.points), [])
    const xs = allPoints.map(point => point.x)
    const ys = allPoints.map(point => point.y)
    const scale = Math.max(
      Math.max.apply(null, xs) - Math.min.apply(null, xs),
      Math.max.apply(null, ys) - Math.min.apply(null, ys),
      1
    )
    const epsilon = scale * 1e-7
    const segments = []
    usablePaths.forEach((path, pathIndex) => {
      const edgeCount = path.closed ? path.points.length : path.points.length - 1
      for (let i = 0; i < edgeCount; i += 1) {
        const a = path.points[i]
        const b = path.points[(i + 1) % path.points.length]
        if (Math.hypot(b.x - a.x, b.y - a.y) <= epsilon) continue
        segments.push({ a, b, pathIndex, splits: [0, 1] })
      }
    })
    if (segments.length > 6000) {
      throw new Error('线稿采样后超过 6000 段，请降低曲边精度或减少参与识别的对象。')
    }

    const bounds = segments.map(segment => ({
      left: Math.min(segment.a.x, segment.b.x) - epsilon,
      right: Math.max(segment.a.x, segment.b.x) + epsilon,
      top: Math.min(segment.a.y, segment.b.y) - epsilon,
      bottom: Math.max(segment.a.y, segment.b.y) + epsilon
    }))
    for (let i = 0; i < segments.length; i += 1) {
      for (let j = i + 1; j < segments.length; j += 1) {
        if (
          bounds[i].right < bounds[j].left || bounds[j].right < bounds[i].left ||
          bounds[i].bottom < bounds[j].top || bounds[j].bottom < bounds[i].top
        ) continue
        const hit = segmentIntersection(segments[i], segments[j], epsilon)
        if (!hit) continue
        if (hit.collinear) addCollinearSegmentSplits(segments[i], segments[j], epsilon)
        else {
          segments[i].splits.push(hit.t)
          segments[j].splits.push(hit.u)
        }
      }
    }

    const vertices = []
    const vertexBuckets = new Map()
    const bucketSize = Math.max(epsilon * 4, 1e-7)
    const bucketKey = (x, y) => x + ':' + y
    const getVertexId = point => {
      const bx = Math.round(point.x / bucketSize)
      const by = Math.round(point.y / bucketSize)
      for (let dx = -1; dx <= 1; dx += 1) {
        for (let dy = -1; dy <= 1; dy += 1) {
          const ids = vertexBuckets.get(bucketKey(bx + dx, by + dy)) || []
          for (const id of ids) {
            if (Math.hypot(vertices[id].x - point.x, vertices[id].y - point.y) <= epsilon * 4) return id
          }
        }
      }
      const id = vertices.length
      vertices.push({ x: point.x, y: point.y })
      const key = bucketKey(bx, by)
      if (!vertexBuckets.has(key)) vertexBuckets.set(key, [])
      vertexBuckets.get(key).push(id)
      return id
    }

    const edges = new Map()
    segments.forEach(segment => {
      const length = Math.hypot(segment.b.x - segment.a.x, segment.b.y - segment.a.y)
      const splitEpsilon = epsilon / Math.max(length, 1)
      const splits = uniqueSortedNumbers(segment.splits.map(value => Math.max(0, Math.min(1, value))), splitEpsilon)
      for (let i = 0; i < splits.length - 1; i += 1) {
        const a = interpolatePoint(segment.a, segment.b, splits[i])
        const b = interpolatePoint(segment.a, segment.b, splits[i + 1])
        if (Math.hypot(b.x - a.x, b.y - a.y) <= epsilon) continue
        const u = getVertexId(a)
        const v = getVertexId(b)
        if (u === v) continue
        const key = u < v ? u + ':' + v : v + ':' + u
        if (!edges.has(key)) edges.set(key, { u, v })
      }
    })

    const regions = enumerateGraphFaces(vertices, Array.from(edges.values()), false)
      .map(region => cleanPolygon(region, epsilon))
      .filter(region => region.length >= 3 && Math.abs(polygonArea(region)) > scale * scale * 1e-8)
    if (!regions.length) {
      throw new Error('通用线稿中没有形成闭合区域；请确认线条端点相接或彼此相交。')
    }
    return regions
  }

  function buildPlanarRegions(boundary, lines) {
    const xs = boundary.map(point => point.x)
    const ys = boundary.map(point => point.y)
    const scale = Math.max(Math.max.apply(null, xs) - Math.min.apply(null, xs), Math.max.apply(null, ys) - Math.min.apply(null, ys), 1)
    const epsilon = scale * 1e-7
    const minArea = Math.max(Math.abs(polygonArea(boundary)) * 1e-8, epsilon * epsilon * 10)
    const segments = []

    for (let i = 0; i < boundary.length; i += 1) {
      const a = boundary[i]
      const b = boundary[(i + 1) % boundary.length]
      if (Math.hypot(b.x - a.x, b.y - a.y) > epsilon) {
        segments.push({ a, b, kind: 'boundary', splits: [0, 1] })
      }
    }
    for (const line of lines) {
      const pieces = dividerSegmentsInsideBoundary(line, boundary, epsilon)
      for (const piece of pieces) segments.push(Object.assign(piece, { splits: [0, 1] }))
    }

    for (let i = 0; i < segments.length; i += 1) {
      for (let j = i + 1; j < segments.length; j += 1) {
        const hit = segmentIntersection(segments[i], segments[j], epsilon)
        if (!hit) continue
        if (hit.collinear) {
          if (segments[i].kind === 'divider' && segments[j].kind === 'divider') {
            throw new Error('检测到重合的分割线，请删除重复或共线重叠的直线。')
          }
          continue
        }
        segments[i].splits.push(hit.t)
        segments[j].splits.push(hit.u)
      }
    }

    const vertices = []
    const getVertexId = point => {
      for (let i = 0; i < vertices.length; i += 1) {
        if (Math.hypot(vertices[i].x - point.x, vertices[i].y - point.y) <= epsilon * 4) return i
      }
      vertices.push({ x: point.x, y: point.y })
      return vertices.length - 1
    }

    const edges = new Map()
    for (const segment of segments) {
      const splitEpsilon = epsilon / Math.max(Math.hypot(segment.b.x - segment.a.x, segment.b.y - segment.a.y), 1)
      const splits = uniqueSortedNumbers(segment.splits, splitEpsilon)
      for (let i = 0; i < splits.length - 1; i += 1) {
        const a = interpolatePoint(segment.a, segment.b, splits[i])
        const b = interpolatePoint(segment.a, segment.b, splits[i + 1])
        if (Math.hypot(b.x - a.x, b.y - a.y) <= epsilon) continue
        const u = getVertexId(a)
        const v = getVertexId(b)
        if (u === v) continue
        const key = u < v ? u + ':' + v : v + ':' + u
        if (!edges.has(key)) edges.set(key, { u, v })
      }
    }

    const adjacency = Array.from({ length: vertices.length }, () => [])
    edges.forEach(edge => {
      adjacency[edge.u].push(edge.v)
      adjacency[edge.v].push(edge.u)
    })
    adjacency.forEach((neighbors, index) => {
      neighbors.sort((a, b) => {
        const angleA = Math.atan2(vertices[a].y - vertices[index].y, vertices[a].x - vertices[index].x)
        const angleB = Math.atan2(vertices[b].y - vertices[index].y, vertices[b].x - vertices[index].x)
        return angleA - angleB
      })
    })

    const visited = new Set()
    const faces = []
    const directedKey = (u, v) => u + '>' + v
    edges.forEach(edge => {
      for (const direction of [[edge.u, edge.v], [edge.v, edge.u]]) {
        let u = direction[0]
        let v = direction[1]
        const startU = u
        const startV = v
        if (visited.has(directedKey(u, v))) continue
        const face = []
        let closed = false
        const guardLimit = edges.size * 2 + 5
        for (let guard = 0; guard < guardLimit; guard += 1) {
          const key = directedKey(u, v)
          if (visited.has(key) && !(u === startU && v === startV)) break
          visited.add(key)
          face.push(vertices[u])
          const neighbors = adjacency[v]
          const reverseIndex = neighbors.indexOf(u)
          if (reverseIndex < 0 || !neighbors.length) break
          const w = neighbors[(reverseIndex - 1 + neighbors.length) % neighbors.length]
          u = v
          v = w
          if (u === startU && v === startV) {
            closed = true
            break
          }
        }
        if (!closed) continue
        const cleaned = cleanPolygon(face, epsilon)
        const area = cleaned.length >= 3 ? polygonArea(cleaned) : 0
        if (area > minArea) faces.push(cleaned)
      }
    })

    const regions = faces.filter(face => {
      const centroid = polygonCentroid(face)
      if (pointInPolygon(centroid, boundary, epsilon)) return true
      for (let i = 0; i < face.length; i += 1) {
        const a = face[i]
        const b = face[(i + 1) % face.length]
        const dx = b.x - a.x
        const dy = b.y - a.y
        const length = Math.hypot(dx, dy)
        if (length <= epsilon) continue
        const probe = {
          x: (a.x + b.x) / 2 - dy / length * epsilon * 20,
          y: (a.y + b.y) / 2 + dx / length * epsilon * 20
        }
        if (pointInPolygon(probe, boundary, epsilon)) return true
      }
      return false
    })

    regions.sort((a, b) => {
      const ca = polygonCentroid(a)
      const cb = polygonCentroid(b)
      return Math.abs(ca.y - cb.y) > 1e-6 ? ca.y - cb.y : ca.x - cb.x
    })
    if (!regions.length) throw new Error('没有识别出可生成的闭合区域。')
    return regions
  }

  function enumerateGraphFaces(vertices, graphEdges, includeNodeIds) {
    if (vertices.length < 3 || graphEdges.length < 3) return []
    const xs = vertices.map(point => point.x)
    const ys = vertices.map(point => point.y)
    const scale = Math.max(
      Math.max.apply(null, xs) - Math.min.apply(null, xs),
      Math.max.apply(null, ys) - Math.min.apply(null, ys),
      1
    )
    const epsilon = scale * 1e-7
    const minArea = scale * scale * 1e-8
    const adjacency = Array.from({ length: vertices.length }, () => [])
    const edgeMap = new Map()

    for (const edge of graphEdges) {
      if (edge.u === edge.v) continue
      const key = edge.u < edge.v ? edge.u + ':' + edge.v : edge.v + ':' + edge.u
      if (edgeMap.has(key)) continue
      edgeMap.set(key, edge)
      adjacency[edge.u].push(edge.v)
      adjacency[edge.v].push(edge.u)
    }
    adjacency.forEach((neighbors, index) => {
      neighbors.sort((a, b) => {
        const angleA = Math.atan2(vertices[a].y - vertices[index].y, vertices[a].x - vertices[index].x)
        const angleB = Math.atan2(vertices[b].y - vertices[index].y, vertices[b].x - vertices[index].x)
        return angleA - angleB
      })
    })

    const visited = new Set()
    const directedKey = (u, v) => u + '>' + v
    const regions = []
    edgeMap.forEach(edge => {
      for (const direction of [[edge.u, edge.v], [edge.v, edge.u]]) {
        let u = direction[0]
        let v = direction[1]
        const startU = u
        const startV = v
        if (visited.has(directedKey(u, v))) continue
        const faceNodeIds = []
        let closed = false
        const guardLimit = edgeMap.size * 2 + 5
        for (let guard = 0; guard < guardLimit; guard += 1) {
          const key = directedKey(u, v)
          if (visited.has(key) && !(u === startU && v === startV)) break
          visited.add(key)
          faceNodeIds.push(u)
          const neighbors = adjacency[v]
          const reverseIndex = neighbors.indexOf(u)
          if (reverseIndex < 0 || !neighbors.length) break
          const w = neighbors[(reverseIndex - 1 + neighbors.length) % neighbors.length]
          u = v
          v = w
          if (u === startU && v === startV) {
            closed = true
            break
          }
        }
        if (!closed) continue
        const face = faceNodeIds.map(index => vertices[index])
        if (face.length >= 3 && polygonArea(face) > minArea) {
          regions.push({ nodeIds: faceNodeIds, points: face })
        }
      }
    })

    regions.sort((a, b) => {
      const ca = polygonCentroid(a.points)
      const cb = polygonCentroid(b.points)
      return Math.abs(ca.y - cb.y) > 1e-6 ? ca.y - cb.y : ca.x - cb.x
    })
    return includeNodeIds ? regions : regions.map(region => region.points)
  }

  function ellipsePointOnRay(ellipse, angle) {
    const dx = Math.cos(angle)
    const dy = Math.sin(angle)
    const scale = 1 / Math.sqrt(
      dx * dx / Math.max(ellipse.rx * ellipse.rx, 1e-12) +
      dy * dy / Math.max(ellipse.ry * ellipse.ry, 1e-12)
    )
    return {
      x: ellipse.cx + dx * scale,
      y: ellipse.cy + dy * scale
    }
  }

  function buildCurvedNetworkFace(nodeIds, nodes, segments, graphEdges) {
    const points = []
    const fullTurn = Math.PI * 2
    const edgeMap = new Map()
    ;(graphEdges || []).forEach(edge => {
      const key = edge.u < edge.v ? edge.u + ':' + edge.v : edge.v + ':' + edge.u
      if (!edgeMap.has(key)) edgeMap.set(key, edge)
    })
    const attachmentAngle = (nodeIndex, neighborIndex) => {
      const node = nodes[nodeIndex]
      const key = nodeIndex < neighborIndex
        ? nodeIndex + ':' + neighborIndex
        : neighborIndex + ':' + nodeIndex
      const edge = edgeMap.get(key)
      let target = null
      if (edge) target = edge.u === nodeIndex ? edge.pointU : edge.pointV
      if (!target || Math.hypot(target.x - node.cx, target.y - node.cy) < 1e-8) {
        const neighbor = nodes[neighborIndex]
        target = { x: neighbor.cx, y: neighbor.cy }
      }
      return Math.atan2(target.y - node.cy, target.x - node.cx)
    }
    for (let i = 0; i < nodeIds.length; i += 1) {
      const previousIndex = nodeIds[(i - 1 + nodeIds.length) % nodeIds.length]
      const currentIndex = nodeIds[i]
      const nextIndex = nodeIds[(i + 1) % nodeIds.length]
      const current = nodes[currentIndex]
      const startAngle = attachmentAngle(currentIndex, previousIndex)
      const endAngle = attachmentAngle(currentIndex, nextIndex)
      let clockwiseGap = (startAngle - endAngle) % fullTurn
      if (clockwiseGap < 0) clockwiseGap += fullTurn
      if (clockwiseGap < 1e-8) clockwiseGap = fullTurn
      const arcSteps = Math.max(2, Math.ceil(segments * clockwiseGap / fullTurn))
      for (let step = 0; step <= arcSteps; step += 1) {
        const angle = startAngle - clockwiseGap * step / arcSteps
        pushDistinct(points, ellipsePointOnRay(current, angle), 1e-6)
      }
    }
    return cleanPolygon(points, 1e-6)
  }

  function buildNetworkModel(nodeShapes, lineShapes, segments, includeNodeInteriors) {
    const nodes = nodeShapes.map(shape => {
      const ellipse = ellipseFromShape(shape)
      return Object.assign(ellipse, { name: shapeName(shape) })
    })
    const lines = lineShapes.map(lineFromShape)

    const matchEndpoint = point => {
      let bestIndex = -1
      let bestScore = Infinity
      for (let i = 0; i < nodes.length; i += 1) {
        const node = nodes[i]
        const nx = (point.x - node.cx) / Math.max(node.rx, 1e-8)
        const ny = (point.y - node.cy) / Math.max(node.ry, 1e-8)
        const normalizedRadius = Math.sqrt(nx * nx + ny * ny)
        const boundaryDistance = Math.abs(normalizedRadius - 1) * Math.min(node.rx, node.ry)
        const centerDistance = Math.hypot(point.x - node.cx, point.y - node.cy)
        const score = Math.min(boundaryDistance, centerDistance)
        const tolerance = Math.max(3, Math.min(node.rx, node.ry) * 1.15)
        if (score <= tolerance && score < bestScore) {
          bestScore = score
          bestIndex = i
        }
      }
      return bestIndex
    }

    const graphEdges = []
    const ignoredLines = []
    for (const line of lines) {
      const u = matchEndpoint(line.p1)
      const v = matchEndpoint(line.p2)
      if (u < 0 || v < 0 || u === v) {
        ignoredLines.push(line.name)
        continue
      }
      graphEdges.push({
        u,
        v,
        name: line.name,
        pointU: { x: line.p1.x, y: line.p1.y },
        pointV: { x: line.p2.x, y: line.p2.y }
      })
    }

    const vertices = nodes.map(node => ({ x: node.cx, y: node.cy }))
    const faceRecords = enumerateGraphFaces(vertices, graphEdges, true)
    if (!faceRecords.length) {
      throw new Error(
        '检测到了椭圆和线段，但没有找到闭合区域。请确认每条线段的两端确实连接到椭圆边缘，且外围线网已经闭合。'
      )
    }
    const faceRegions = faceRecords.map(face => buildCurvedNetworkFace(face.nodeIds, nodes, segments, graphEdges))
    const nodeRegions = includeNodeInteriors
      ? nodes.map(node => sampleEllipse(node, segments))
      : []
    const regions = faceRegions.concat(nodeRegions)
    return {
      nodes,
      lines,
      graphEdges,
      regions,
      faceRegions,
      nodeRegions,
      ignoredLines
    }
  }

  function buildRegions(boundaryOrEllipse, lines, segments) {
    const boundary = boundaryOrEllipse && Array.isArray(boundaryOrEllipse.points)
      ? boundaryOrEllipse.points
      : boundaryOrEllipse && Array.isArray(boundaryOrEllipse)
        ? boundaryOrEllipse
        : sampleEllipse(boundaryOrEllipse, segments)
    return buildPlanarRegions(cleanPolygon(boundary, 1e-8), lines || [])
  }

  function hexToRgb(hex) {
    const normalized = String(hex || '').trim().replace(/^#/, '')
    if (!/^[0-9a-fA-F]{6}$/.test(normalized)) throw new Error('颜色格式无效：' + hex)
    return {
      r: parseInt(normalized.slice(0, 2), 16),
      g: parseInt(normalized.slice(2, 4), 16),
      b: parseInt(normalized.slice(4, 6), 16)
    }
  }

  function officeRgb(hex) {
    const c = hexToRgb(hex)
    return c.r + c.g * 256 + c.b * 65536
  }

  function rgbToHsl(rgb) {
    let r = rgb.r / 255
    let g = rgb.g / 255
    let b = rgb.b / 255
    const max = Math.max(r, g, b)
    const min = Math.min(r, g, b)
    let h = 0
    let s = 0
    const l = (max + min) / 2
    const d = max - min
    if (d !== 0) {
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6
      else if (max === g) h = ((b - r) / d + 2) / 6
      else h = ((r - g) / d + 4) / 6
    }
    return { h, s, l }
  }

  function hslToHex(hsl) {
    let r
    let g
    let b
    const h = ((hsl.h % 1) + 1) % 1
    const s = Math.max(0, Math.min(1, hsl.s))
    const l = Math.max(0, Math.min(1, hsl.l))
    if (s === 0) r = g = b = l
    else {
      const hue2rgb = (p, q, t0) => {
        let t = t0
        if (t < 0) t += 1
        if (t > 1) t -= 1
        if (t < 1 / 6) return p + (q - p) * 6 * t
        if (t < 1 / 2) return q
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
        return p
      }
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s
      const p = 2 * l - q
      r = hue2rgb(p, q, h + 1 / 3)
      g = hue2rgb(p, q, h)
      b = hue2rgb(p, q, h - 1 / 3)
    }
    const byte = value => Math.round(value * 255).toString(16).padStart(2, '0')
    return '#' + byte(r) + byte(g) + byte(b)
  }

  function paletteColor(baseHex, index, count) {
    if (count <= 1) return baseHex
    const hsl = rgbToHsl(hexToRgb(baseHex))
    const hueOffset = (index / count) * 0.22
    const lightOffset = ((index % 3) - 1) * 0.09
    return hslToHex({
      h: hsl.h + hueOffset,
      s: Math.max(0.35, hsl.s),
      l: Math.max(0.25, Math.min(0.78, hsl.l + lightOffset))
    })
  }

  function setFill(shape, colorHex, opacityPercent) {
    shape.Fill.Visible = M.msoTrue
    try { shape.Fill.Solid() } catch (_) {}
    shape.Fill.ForeColor.RGB = officeRgb(colorHex)
    shape.Fill.Transparency = 1 - Math.max(0, Math.min(100, opacityPercent)) / 100
    shape.Line.Visible = M.msoFalse
  }

  function createFreeform(slide, polygon, colorHex, opacityPercent, name) {
    if (!polygon || polygon.length < 3 || polygon.some(point => !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
      throw new Error('检测到无效的区域坐标，已停止生成。')
    }
    const xs = polygon.map(point => point.x)
    const ys = polygon.map(point => point.y)
    const expected = {
      left: Math.min.apply(null, xs),
      top: Math.min.apply(null, ys),
      right: Math.max.apply(null, xs),
      bottom: Math.max.apply(null, ys)
    }
    const first = polygon[0]
    const closedPoints = polygon.concat([{ x: first.x, y: first.y }])
    let shape = null

    // WPS 的 BuildFreeform/AddNodes 在不同版本中对可选坐标参数的解释不一致。
    // AddPolyline 直接接收二维点数组，坐标最稳定，并且首尾同点时可形成可填充闭合区域。
    try {
      if (slide.Shapes && typeof slide.Shapes.AddPolyline === 'function') {
        shape = slide.Shapes.AddPolyline(closedPoints.map(point => [point.x, point.y]))
      }
    } catch (_) {
      shape = null
    }

    if (!shape) {
      const builder = slide.Shapes.BuildFreeform(M.msoEditingCorner, first.x, first.y)
      for (let i = 1; i < polygon.length; i += 1) {
        const point = polygon[i]
        builder.AddNodes(
          M.msoSegmentLine,
          M.msoEditingCorner,
          point.x,
          point.y,
          point.x,
          point.y,
          point.x,
          point.y
        )
      }
      builder.AddNodes(
        M.msoSegmentLine,
        M.msoEditingCorner,
        first.x,
        first.y,
        first.x,
        first.y,
        first.x,
        first.y
      )
      shape = builder.ConvertToShape()
    }
    const actual = {
      left: toNumber(shape.Left, expected.left),
      top: toNumber(shape.Top, expected.top),
      right: toNumber(shape.Left, expected.left) + Math.abs(toNumber(shape.Width, expected.right - expected.left)),
      bottom: toNumber(shape.Top, expected.top) + Math.abs(toNumber(shape.Height, expected.bottom - expected.top))
    }
    const expectedSize = Math.max(expected.right - expected.left, expected.bottom - expected.top, 1)
    const boundsTolerance = Math.max(10, expectedSize * 0.25)
    if (
      actual.left < expected.left - boundsTolerance ||
      actual.top < expected.top - boundsTolerance ||
      actual.right > expected.right + boundsTolerance ||
      actual.bottom > expected.bottom + boundsTolerance
    ) {
      try { shape.Delete() } catch (_) {}
      throw new Error('WPS 返回的自由形状坐标超出目标区域，已停止生成以避免产生错位对象。')
    }
    shape.Name = name
    setFill(shape, colorHex, opacityPercent)
    shape.ZOrder(M.msoSendToBack)
    return shape
  }

  function collectSelectionModel(options) {
    const app = getApplication()
    const slide = getActiveSlide(app)
    const selectedTopLevel = getSelectedShapes(app, '', true)
    const selected = expandShapes(selectedTopLevel).filter(shape => !isPluginShape(shape))
    const slideShapes = getSlideShapes(slide)
    let sourceShapes = selected.length ? selected : slideShapes

    // 外框模式的常用流程只需选中外框；插件自动收集当前页直线，
    // 避免普通用户为了“外框 + 全部分割线”必须一直按住 Ctrl。
    if (
      selected.length === 1 &&
      !isPluginShape(selected[0]) &&
      supportedBoundaryKind(selected[0])
    ) {
      sourceShapes = [selected[0]].concat(slideShapes.filter(shape => isStraightLine(shape)))
    }
    const boundaryShapes = sourceShapes.filter(shape => supportedBoundaryKind(shape))
    const ellipseShapes = sourceShapes.filter(shape => isEllipse(shape) && !isPluginShape(shape))
    const lineShapes = sourceShapes.filter(shape => isStraightLine(shape) && !isPluginShape(shape))
    const segments = Math.max(48, Math.min(360, Math.round(toNumber(options && options.segments, 160))))
    const includeNodeInteriors = !options || options.includeNodeInteriors !== false

    if (ellipseShapes.length >= 3 && lineShapes.length >= 3 && boundaryShapes.length !== 1) {
      try {
        const network = buildNetworkModel(ellipseShapes, lineShapes, segments, includeNodeInteriors)
        return {
          app,
          slide,
          mode: 'network',
          boundary: null,
          lines: network.lines,
          regions: network.regions,
          nodes: network.nodes,
          paths: [],
          closedShapes: [],
          edgeCount: network.graphEdges.length,
          faceCount: network.faceRegions.length,
          nodeRegionCount: network.nodeRegions.length,
          ignoredLines: network.ignoredLines,
          scannedWholeSlide: selectedTopLevel.length === 0,
          selectedCount: selected.length
        }
      } catch (_) {
        // 不满足椭圆节点线网连接规则时，继续尝试通用线稿模式。
      }
    }

    if (boundaryShapes.length === 1) {
      try {
        const boundary = boundaryFromShape(boundaryShapes[0], segments)
        const lines = lineShapes.map(lineFromShape)
        const regions = buildRegions(boundary, lines, segments)
        return {
          app,
          slide,
          mode: 'boundary',
          boundary,
          lines,
          regions,
          nodes: [],
          paths: [],
          closedShapes: [boundary.shape],
          edgeCount: lines.length,
          faceCount: regions.length,
          nodeRegionCount: 0,
          ignoredLines: [],
          scannedWholeSlide: selectedTopLevel.length === 0,
          selectedCount: selected.length
        }
      } catch (error) {
        // 开放自由曲线不能作为单一外框；让通用线稿模式继续尝试。
        if (getShapeType(boundaryShapes[0]) !== M.msoFreeform) throw error
      }
    }

    const pathFailures = []
    const paths = sourceShapes.map(shape => {
      try {
        const path = genericPathFromShape(shape, segments)
        if (!path) pathFailures.push(shapeName(shape) + '（不支持的对象类型）')
        return path
      } catch (error) {
        pathFailures.push(shapeName(shape) + '（' + (error && error.message ? error.message : '读取失败') + '）')
        return null
      }
    }).filter(Boolean)
    let genericRegions
    try {
      genericRegions = buildGenericRegions(paths)
    } catch (error) {
      const failed = pathFailures.length ? '；未读取：' + pathFailures.slice(0, 3).join('、') : ''
      throw new Error(
        '通用线稿已扫描 ' + sourceShapes.length + ' 个对象，成功读取 ' + paths.length + ' 条路径。' +
        (error && error.message ? error.message : String(error)) + failed
      )
    }
    const closedShapes = paths.filter(path => path.closed).map(path => path.shape)
    return {
      app,
      slide,
      mode: 'generic',
      boundary: null,
      lines: [],
      regions: genericRegions,
      nodes: [],
      paths,
      closedShapes,
      edgeCount: paths.reduce((sum, path) => sum + Math.max(0, path.points.length - (path.closed ? 0 : 1)), 0),
      faceCount: genericRegions.length,
      nodeRegionCount: 0,
      ignoredLines: [],
      scannedWholeSlide: selectedTopLevel.length === 0,
      selectedCount: selected.length
    }
  }

  function analyze(options) {
    const model = collectSelectionModel(options || {})
    if (model.mode === 'network') {
      return {
        mode: model.mode,
        boundaryName: '当前页线网',
        boundaryKind: '椭圆节点网络',
        nodeCount: model.nodes.length,
        lineCount: model.edgeCount,
        ignoredLineCount: model.ignoredLines.length,
        faceCount: model.faceCount,
        nodeRegionCount: model.nodeRegionCount,
        regionCount: model.regions.length,
        vertexCount: model.regions.reduce((sum, poly) => sum + poly.length, 0),
        scannedWholeSlide: model.scannedWholeSlide
      }
    }
    if (model.mode === 'generic') {
      return {
        mode: model.mode,
        boundaryName: '当前线稿',
        boundaryKind: '通用线稿',
        nodeCount: 0,
        lineCount: model.paths.length,
        ignoredLineCount: 0,
        faceCount: model.faceCount,
        nodeRegionCount: 0,
        regionCount: model.regions.length,
        vertexCount: model.regions.reduce((sum, poly) => sum + poly.length, 0),
        scannedWholeSlide: model.scannedWholeSlide
      }
    }
    return {
      mode: model.mode,
      boundaryName: model.boundary.name,
      boundaryKind: model.boundary.kind,
      nodeCount: 0,
      lineCount: model.lines.length,
      ignoredLineCount: 0,
      faceCount: model.regions.length,
      nodeRegionCount: 0,
      regionCount: model.regions.length,
      vertexCount: model.regions.reduce((sum, poly) => sum + poly.length, 0),
      scannedWholeSlide: model.scannedWholeSlide
    }
  }

  function clearPreviewShapes(slide) {
    let deleted = 0
    for (let i = toNumber(slide.Shapes.Count, 0); i >= 1; i -= 1) {
      const shape = slide.Shapes.Item(i)
      if (isPreviewShape(shape)) {
        shape.Delete()
        deleted += 1
      }
    }
    return deleted
  }

  function previewSelectionKey(shape) {
    return shapeName(shape)
      .replace(PICKED_PREVIEW_PREFIX, '')
      .replace(PREVIEW_PREFIX, '')
  }

  function countPickedPreviews(slide) {
    let count = 0
    for (let i = 1; i <= toNumber(slide.Shapes.Count, 0); i += 1) {
      if (isPickedPreviewShape(slide.Shapes.Item(i))) count += 1
    }
    return count
  }

  function setPreviewPicked(shape, picked, color) {
    const key = previewSelectionKey(shape)
    shape.Name = (picked ? PICKED_PREVIEW_PREFIX : PREVIEW_PREFIX) + key
    setFill(shape, color, picked ? 46 : 18)
  }

  function captureManualSelection(options) {
    const settings = Object.assign({ color: '#4f7cff' }, options || {})
    const app = getApplication()
    const slide = getActiveSlide(app)
    const selected = getSelectedShapes(app, '', true).filter(isPreviewShape)

    if (!selected.length) {
      lastPreviewSelectionKey = ''
      return { changed: false, count: countPickedPreviews(slide) }
    }

    const selectionKey = selected.map(previewSelectionKey).sort().join('|')
    if (selectionKey === lastPreviewSelectionKey) {
      unselectCurrentShapes(app)
      return { changed: false, count: countPickedPreviews(slide) }
    }

    lastPreviewSelectionKey = selectionKey
    selected.forEach(shape => {
      setPreviewPicked(shape, !isPickedPreviewShape(shape), settings.color)
    })
    unselectCurrentShapes(app)
    return {
      changed: true,
      count: countPickedPreviews(slide)
    }
  }

  function disableBoundaryFill(model, settings) {
    const disable = settings.disableBoundaryFill !== undefined
      ? settings.disableBoundaryFill
      : settings.disableEllipseFill
    if (model.mode === 'network') {
      if (disable) {
        model.nodes.forEach(node => {
          try { node.shape.Fill.Visible = M.msoFalse } catch (_) {}
        })
      }
      return
    }
    if (model.mode === 'generic') {
      if (disable) {
        model.closedShapes.forEach(shape => {
          try { shape.Fill.Visible = M.msoFalse } catch (_) {}
        })
      }
      return
    }
    if (!model.boundary) return
    if (disable) {
      try { model.boundary.shape.Fill.Visible = M.msoFalse } catch (_) {}
    }
  }

  function generationSummary(model, names) {
    return {
      count: names.length,
      names,
      mode: model.mode,
      faceCount: model.faceCount,
      nodeRegionCount: model.nodeRegionCount,
      ignoredLineCount: model.ignoredLines.length
    }
  }

  function generate(options) {
    const settings = Object.assign({
      color: '#4f7cff',
      opacity: 72,
      segments: 160,
      palette: false,
      disableBoundaryFill: true
    }, options || {})

    const model = collectSelectionModel(settings)
    clearPreviewShapes(model.slide)
    disableBoundaryFill(model, settings)

    const runId = Date.now().toString(36)
    const names = []
    for (let i = 0; i < model.regions.length; i += 1) {
      const color = settings.palette
        ? paletteColor(settings.color, i, model.regions.length)
        : settings.color
      const name = PREFIX + runId + '_' + String(i + 1).padStart(2, '0')
      createFreeform(model.slide, model.regions[i], color, settings.opacity, name)
      names.push(name)
    }
    unselectCurrentShapes(model.app)
    return generationSummary(model, names)
  }

  function prepareManualSelection(options) {
    const settings = Object.assign({
      color: '#4f7cff',
      segments: 160
    }, options || {})
    const model = collectSelectionModel(settings)
    clearPreviewShapes(model.slide)
    lastPreviewSelectionKey = ''
    disableBoundaryFill(model, settings)

    const runId = Date.now().toString(36)
    const names = []
    for (let i = 0; i < model.regions.length; i += 1) {
      const name = PREVIEW_PREFIX + runId + '_' + String(i + 1).padStart(2, '0')
      createFreeform(model.slide, model.regions[i], settings.color, 18, name)
      names.push(name)
    }
    unselectCurrentShapes(model.app)
    return generationSummary(model, names)
  }

  function generateSelected(options) {
    const settings = Object.assign({
      color: '#4f7cff',
      opacity: 72,
      palette: false
    }, options || {})
    const app = getApplication()
    const slide = getActiveSlide(app)

    // 若用户刚点完区域便立即点击任务窗格，轮询可能还未来得及记录；
    // 在提交前再收集一次当前 WPS 选择，确保这次点击不会丢失。
    const currentlySelected = getSelectedShapes(app, '', true).filter(isPreviewShape)
    currentlySelected.forEach(shape => {
      if (!isPickedPreviewShape(shape)) setPreviewPicked(shape, true, settings.color)
    })
    const previews = []
    for (let i = 1; i <= toNumber(slide.Shapes.Count, 0); i += 1) {
      const shape = slide.Shapes.Item(i)
      if (isPickedPreviewShape(shape)) previews.push(shape)
    }
    if (!previews.length) {
      throw new Error('还没有点选候选区域。请直接点击一个或多个浅色区域，已选区域会自动加深。')
    }

    const runId = Date.now().toString(36)
    const names = []
    for (let i = 0; i < previews.length; i += 1) {
      const shape = previews[i]
      const color = settings.palette
        ? paletteColor(settings.color, i, previews.length)
        : settings.color
      const name = PREFIX + runId + '_' + String(i + 1).padStart(2, '0')
      shape.Name = name
      setFill(shape, color, settings.opacity)
      names.push(name)
    }

    for (let i = toNumber(slide.Shapes.Count, 0); i >= 1; i -= 1) {
      const shape = slide.Shapes.Item(i)
      if (isPreviewShape(shape)) shape.Delete()
    }
    unselectCurrentShapes(app)
    lastPreviewSelectionKey = ''
    return { count: names.length, names }
  }

  function recolorSelected(options) {
    const settings = Object.assign({ color: '#4f7cff', opacity: 72 }, options || {})
    const app = getApplication()
    const selected = getSelectedShapes(app, '请先选中一个或多个已生成的填色区域。')
    const generated = selected.filter(isGeneratedShape)
    if (!generated.length) throw new Error('请先选中一个或多个由本插件生成的填色区域。')
    generated.forEach(shape => setFill(shape, settings.color, settings.opacity))
    return generated.length
  }

  function deleteGenerated() {
    const app = getApplication()
    const slide = getActiveSlide(app)
    let deleted = 0
    for (let i = toNumber(slide.Shapes.Count, 0); i >= 1; i -= 1) {
      const shape = slide.Shapes.Item(i)
      if (isPluginShape(shape)) {
        shape.Delete()
        deleted += 1
      }
    }
    return deleted
  }

  const api = {
    PREFIX,
    PREVIEW_PREFIX,
    PICKED_PREVIEW_PREFIX,
    analyze,
    generate,
    prepareManualSelection,
    captureManualSelection,
    generateSelected,
    recolorSelected,
    deleteGenerated,
    _geometry: {
      sampleEllipse,
      splitPolygon,
      buildRegions,
      buildPlanarRegions,
      buildGenericRegions,
      enumerateGraphFaces,
      buildCurvedNetworkFace,
      sampleFreeformPath,
      ellipsePointOnRay,
      polygonArea,
      polygonCentroid,
      segmentEllipseRoots,
      pointInPolygon
    }
  }

  global.WpsAutoFill = api
  if (typeof module !== 'undefined' && module.exports) module.exports = api
})(typeof window !== 'undefined' ? window : globalThis)
