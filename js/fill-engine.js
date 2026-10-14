(function (global) {
  'use strict'

  const PREFIX = 'WPS_AutoFill_'
  const PREVIEW_PREFIX = PREFIX + 'Preview_'
  const PICKED_PREVIEW_PREFIX = PREVIEW_PREFIX + 'Picked_'
  const MARKER_PREFIX = PREFIX + 'Marker_'
  const PICKED_MARKER_PREFIX = MARKER_PREFIX + 'Picked_'
  const DEBUG_PREFIX = PREFIX + 'GeometryDebug_'
  let lastPreviewSelectionKey = ''
  let geometrySnapshot = null
  const M = {
    msoFalse: 0,
    msoTrue: -1,
    msoAutoShape: 1,
    msoFreeform: 5,
    msoGroup: 6,
    msoEditingAuto: 0,
    msoEditingCorner: 1,
    msoEditingSmooth: 2,
    msoSegmentLine: 0,
    msoSegmentCurve: 1,
    msoBringToFront: 0,
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
    msoShapeSnip1Rectangle: 155,
    msoShapeSnip2SameRectangle: 156,
    msoShapeHalfFrame: 159,
    msoShapeTypeLine: 9,
    msoConnectorStraight: 1,
    msoConnectorElbow: 2,
    msoConnectorCurve: 3
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

  async function refreshGeometrySnapshot() {
    if (!global.WpsOoxmlGeometry || typeof global.WpsOoxmlGeometry.refreshFromApplication !== 'function') {
      throw new Error('真实 PPTX 几何读取组件尚未加载。')
    }
    geometrySnapshot = await global.WpsOoxmlGeometry.refreshFromApplication(getApplication())
    return {
      shapeCount: geometrySnapshot && geometrySnapshot.shapes
        ? Object.keys(geometrySnapshot.shapes).length
        : 0,
      slidePart: geometrySnapshot && geometrySnapshot.partPath
    }
  }

  function setGeometrySnapshot(snapshot) {
    geometrySnapshot = snapshot || null
  }

  function ooxmlPathsForShape(shape, segments) {
    if (!geometrySnapshot || typeof geometrySnapshot.geometryForShape !== 'function') return []
    try {
      return geometrySnapshot.geometryForShape(shape, segments || 160) || []
    } catch (_) {
      return []
    }
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
    return name.indexOf(PREFIX) === 0 &&
      name.indexOf(PREVIEW_PREFIX) !== 0 &&
      name.indexOf(MARKER_PREFIX) !== 0
  }

  function isPreviewShape(shape) {
    return shapeName(shape).indexOf(PREVIEW_PREFIX) === 0
  }

  function isPickedPreviewShape(shape) {
    return shapeName(shape).indexOf(PICKED_PREVIEW_PREFIX) === 0
  }

  function isMarkerShape(shape) {
    return shapeName(shape).indexOf(MARKER_PREFIX) === 0
  }

  function isPickedMarkerShape(shape) {
    return shapeName(shape).indexOf(PICKED_MARKER_PREFIX) === 0
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

    try {
      if (isTrue(shape.Connector)) {
        const connectorType = Number(shape.ConnectorFormat.Type)
        return connectorType === M.msoConnectorStraight
      }
    } catch (_) {
      // 普通直线通常没有可用的 ConnectorFormat。
    }
    return type === M.msoShapeTypeLine
  }

  function getConnectorType(shape) {
    try { return Number(shape.ConnectorFormat.Type) } catch (_) { return NaN }
  }

  function isCurvedConnector(shape) {
    return getConnectorType(shape) === M.msoConnectorCurve || /曲线连接符|curved connector/i.test(shapeName(shape))
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
      case M.msoShapeSnip1Rectangle:
        return [[0, 0], [0.83333, 0], [1, 0.16667], [1, 1], [0, 1]]
      case M.msoShapeSnip2SameRectangle:
        return [[0.16667, 0], [0.83333, 0], [1, 0.16667], [1, 1], [0, 1]]
      case M.msoShapeHalfFrame:
        return [[0, 0], [1, 0], [0.66667, 0.33333], [0.33333, 0.33333], [0.33333, 0.66667], [0, 1]]
      default:
        return null
    }
  }

  function shapeAdjustment(shape, index, fallback) {
    let value = NaN
    try { value = Number(shape.Adjustments.Item(index)) } catch (_) {}
    if (!Number.isFinite(value)) return fallback
    // WPS 12.1 的 JS 桥接层会对不同预设形状统一返回数字 1（实际是无效
    // 的 Item 包装结果），不能把它当成真实调节点比例。
    if (Math.abs(value - 1) <= 1e-12) return fallback
    if (Math.abs(value) > 1) value /= 100000
    return Math.max(0, Math.min(0.5, value))
  }

  function curvedConnectorAdjustment(shape) {
    let value = NaN
    try { value = Number(shape.Adjustments.Item(1)) } catch (_) {}
    if (!Number.isFinite(value) || Math.abs(value - 1) <= 1e-12) return 0.5
    if (Math.abs(value) > 10) value /= 100000
    return Math.max(-2, Math.min(3, value))
  }

  function presetCurvedConnector3Vertices(shape) {
    const bounds = shapeBounds(shape)
    const transform = shapeTransform(shape, bounds)
    const x2 = curvedConnectorAdjustment(shape)
    const x1 = x2 / 2
    const x3 = (1 + x2) / 2
    // ECMA-376 presetShapeDefinitions.xml / curvedConnector3：
    // M(0,0) C(x1,0 x2,1/4 x2,1/2) C(x2,3/4 x3,1 1,1)
    return [
      [0, 0],
      [x1, 0],
      [x2, 0.25],
      [x2, 0.5],
      [x2, 0.75],
      [x3, 1],
      [1, 1]
    ].map(point => transform(point[0], point[1]))
  }

  function adjustedNormalizedPolygon(shape, bounds) {
    const type = getAutoShapeType(shape)
    const shortSide = Math.min(bounds.width, bounds.height)
    if (type === M.msoShapeSnip1Rectangle) {
      const adjustment = shapeAdjustment(shape, 1, 1 / 6)
      const cut = shortSide * adjustment
      return [[0, 0], [1 - cut / bounds.width, 0], [1, cut / bounds.height], [1, 1], [0, 1]]
    }
    if (type === M.msoShapeSnip2SameRectangle) {
      const topAdjustment = shapeAdjustment(shape, 1, 1 / 6)
      const bottomAdjustment = shapeAdjustment(shape, 2, 0)
      const topCutX = shortSide * topAdjustment / bounds.width
      const topCutY = shortSide * topAdjustment / bounds.height
      const bottomCutX = shortSide * bottomAdjustment / bounds.width
      const bottomCutY = shortSide * bottomAdjustment / bounds.height
      return [
        [topCutX, 0], [1 - topCutX, 0], [1, topCutY],
        [1, 1 - bottomCutY], [1 - bottomCutX, 1], [bottomCutX, 1],
        [0, 1 - bottomCutY], [0, topCutY]
      ]
    }
    if (type === M.msoShapeHalfFrame) {
      const verticalAdjustment = shapeAdjustment(shape, 1, 1 / 3)
      const horizontalAdjustment = shapeAdjustment(shape, 2, 1 / 3)
      const x1 = shortSide * horizontalAdjustment / bounds.width
      const y1 = shortSide * verticalAdjustment / bounds.height
      // ECMA-376 halfFrame 的交叉换算为：
      // dx2 = y1 * w / h，dy2 = x1 * h / w，故归一化后
      // x2 = 1 - y1、y2 = 1 - x1。宽高不等时不能做成对称缩进。
      const x2 = 1 - y1
      const y2 = 1 - x1
      return [[0, 0], [1, 0], [x2, y1], [x1, y1], [x1, y2], [0, 1]]
    }
    return normalizedPolygon(type)
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

  function sampleCubicVertices(vertices, segments, epsilon) {
    return sampleCubicPath(vertices, segments, epsilon).points
  }

  function sampleCubicPath(vertices, segments, epsilon) {
    if (!vertices || vertices.length < 4 || (vertices.length - 1) % 3 !== 0) {
      return { points: [], edgeMeta: [] }
    }
    const points = []
    const edgeMeta = []
    const curveCount = (vertices.length - 1) / 3
    // 让“快速 / 标准 / 精细”真正控制整条曲线的采样密度。
    // 旧上限 48 会使 160、240 两档在单段曲线上完全失效，生成块能看出折线。
    const steps = Math.max(8, Math.min(240, Math.ceil(toNumber(segments, 160) / Math.max(curveCount, 1))))
    pushDistinct(points, vertices[0], epsilon)
    for (let curve = 0; curve < curveCount; curve += 1) {
      const offset = curve * 3
      const cubic = {
        p0: vertices[offset],
        p1: vertices[offset + 1],
        p2: vertices[offset + 2],
        p3: vertices[offset + 3]
      }
      for (let step = 1; step <= steps; step += 1) {
        const before = points.length
        pushDistinct(points, cubicBezierPoint(cubic.p0, cubic.p1, cubic.p2, cubic.p3, step / steps), epsilon)
        if (points.length > before) {
          edgeMeta.push({
            kind: 'curve',
            curve: cubic,
            t0: (step - 1) / steps,
            t1: step / steps
          })
        }
      }
    }
    return { points, edgeMeta }
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
    let edgeMeta = []
    if (hasCurve && vertices.length >= 4 && (vertices.length - 1) % 3 === 0) {
      const sampled = sampleCubicPath(vertices, segments, epsilon)
      points = sampled.points
      edgeMeta = sampled.edgeMeta
    } else {
      points = vertices.slice()
    }

    if (closed && points.length > 1) points.pop()
    points = closed && !edgeMeta.length ? cleanPolygon(points, epsilon) : points.reduce((result, point) => {
      pushDistinct(result, point, epsilon)
      return result
    }, [])
    if (edgeMeta.length && edgeMeta.length !== (closed ? points.length : points.length - 1)) edgeMeta = []
    return { shape, name: shapeName(shape), points, closed, curved: hasCurve, edgeMeta }
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
    const ooxmlPaths = ooxmlPathsForShape(shape, 64).filter(path => path.stroke !== false)
    const closedOoxmlPaths = ooxmlPaths.filter(path => path.closed && path.points && path.points.length >= 3)
    // 只有单一闭合轮廓才进入“外框 + 自动收集直线”模式。带孔洞、内部线
    // 或多个子路径的预设形状必须交给通用线稿模式，否则会丢掉内部边界。
    if (closedOoxmlPaths.length === 1 && ooxmlPaths.length === 1) return 'OOXML形状'
    if (ooxmlPaths.length) return ''
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
    let sourcePath = null
    if (kind === 'OOXML形状') {
      const paths = ooxmlPathsForShape(shape, segments)
        .filter(path => path.closed && path.points && path.points.length >= 3)
        .sort((a, b) => Math.abs(polygonArea(b.points)) - Math.abs(polygonArea(a.points)))
      if (!paths.length) throw new Error('PPTX 中没有可用的闭合轮廓：' + shapeName(shape))
      points = paths[0].points
      sourcePath = paths[0]
    } else if (kind === '自由形状') {
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
        points = adjustedNormalizedPolygon(shape, bounds).map(point => transform(point[0], point[1]))
      }
    }

    const area = Math.abs(polygonArea(points))
    if (area <= 1e-6) throw new Error('闭合外框面积过小：' + shapeName(shape))
    return { shape, name: shapeName(shape), kind, points, area, sourcePath }
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

  function sampleCurvedConnectorPath(shape, segments) {
    let vertices = []
    try { vertices = pointPairsFromValue(shape.Vertices) } catch (_) {}

    const line = lineFromShape(shape)
    const scale = Math.max(
      Math.abs(line.p2.x - line.p1.x),
      Math.abs(line.p2.y - line.p1.y),
      1
    )
    const epsilon = scale * 1e-6
    const expectedBounds = {
      left: toNumber(shape.Left, Math.min(line.p1.x, line.p2.x)),
      top: toNumber(shape.Top, Math.min(line.p1.y, line.p2.y)),
      width: Math.abs(toNumber(shape.Width, line.p2.x - line.p1.x)),
      height: Math.abs(toNumber(shape.Height, line.p2.y - line.p1.y))
    }
    const endpointError = candidate => Math.min(
      Math.hypot(candidate[0].x - line.p1.x, candidate[0].y - line.p1.y) +
        Math.hypot(candidate[candidate.length - 1].x - line.p2.x, candidate[candidate.length - 1].y - line.p2.y),
      Math.hypot(candidate[0].x - line.p2.x, candidate[0].y - line.p2.y) +
        Math.hypot(candidate[candidate.length - 1].x - line.p1.x, candidate[candidate.length - 1].y - line.p1.y)
    )
    const candidates = []
    if (vertices.length >= 4 && (vertices.length - 1) % 3 === 0) {
      candidates.push({ source: 'Vertices:absolute', points: vertices })
      candidates.push({ source: 'Vertices:local', points: vertices.map(point => ({
        x: expectedBounds.left + point.x,
        y: expectedBounds.top + point.y
      })) })
      candidates.push({ source: 'Vertices:normalized', points: vertices.map(point => ({
        x: expectedBounds.left + point.x * expectedBounds.width,
        y: expectedBounds.top + point.y * expectedBounds.height
      })) })
    }
    const endpointTolerance = Math.max(2, scale * 0.025)
    let best = null
    let curveSource = ''
    let bestError = Infinity
    candidates.forEach(candidate => {
      const error = endpointError(candidate.points)
      if (error < bestError) {
        best = candidate.points
        curveSource = candidate.source
        bestError = error
      }
    })
    if (!best || bestError > endpointTolerance * 2) {
      curveSource = 'preset:curvedConnector3'
      best = presetCurvedConnector3Vertices(shape)
    } else if (
      Math.hypot(best[0].x - line.p2.x, best[0].y - line.p2.y) <
      Math.hypot(best[0].x - line.p1.x, best[0].y - line.p1.y)
    ) {
      const reversed = []
      for (let i = best.length - 1; i >= 0; i -= 3) {
        if (!reversed.length) reversed.push(best[i])
        if (i >= 3) reversed.push(best[i - 1], best[i - 2], best[i - 3])
      }
      best = reversed
    }
    const sampled = sampleCubicPath(best, segments, epsilon)
    return {
      shape,
      name: shapeName(shape),
      points: sampled.points,
      closed: false,
      curved: true,
      edgeMeta: sampled.edgeMeta,
      curveSource
    }
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

  function distanceToSegment(point, a, b) {
    const dx = b.x - a.x
    const dy = b.y - a.y
    const length2 = dx * dx + dy * dy
    if (length2 <= 1e-16) return Math.hypot(point.x - a.x, point.y - a.y)
    const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / length2))
    return Math.hypot(point.x - (a.x + dx * t), point.y - (a.y + dy * t))
  }

  function polygonInteriorPoint(polygon) {
    const rings = [polygon].concat(polygon._holes || [])
    const contains = (point, epsilon) => pointInPolygon(point, polygon, epsilon) &&
      !(polygon._holes || []).some(hole => pointInPolygon(point, hole, epsilon))
    const xs = polygon.map(point => point.x)
    const ys = polygon.map(point => point.y)
    const bounds = {
      left: Math.min.apply(null, xs),
      right: Math.max.apply(null, xs),
      top: Math.min.apply(null, ys),
      bottom: Math.max.apply(null, ys)
    }
    const scale = Math.max(bounds.right - bounds.left, bounds.bottom - bounds.top, 1)
    const epsilon = scale * 1e-7
    const candidates = []
    const centroid = polygonCentroid(polygon)
    if (contains(centroid, epsilon)) candidates.push(centroid)

    const scanYs = [0.12, 0.22, 0.32, 0.42, 0.5, 0.58, 0.68, 0.78, 0.88]
      .map(ratio => bounds.top + (bounds.bottom - bounds.top) * ratio + epsilon * 3)
    scanYs.forEach(y => {
      const intersections = []
      rings.forEach(ring => { for (let i = 0; i < ring.length; i += 1) {
        const a = ring[i]
        const b = ring[(i + 1) % ring.length]
        if ((a.y > y) === (b.y > y)) continue
        intersections.push(a.x + (b.x - a.x) * (y - a.y) / (b.y - a.y))
      } })
      intersections.sort((a, b) => a - b)
      for (let i = 0; i + 1 < intersections.length; i += 2) {
        if (intersections[i + 1] - intersections[i] > epsilon * 4) {
          candidates.push({ x: (intersections[i] + intersections[i + 1]) / 2, y })
        }
      }
    })

    let best = candidates[0] || polygon[0]
    let bestClearance = -1
    candidates.forEach(candidate => {
      if (!contains(candidate, epsilon)) return
      let clearance = Infinity
      rings.forEach(ring => { for (let i = 0; i < ring.length; i += 1) {
        clearance = Math.min(clearance, distanceToSegment(
          candidate,
          ring[i],
          ring[(i + 1) % ring.length]
        ))
      } })
      if (clearance > bestClearance) {
        best = candidate
        bestClearance = clearance
      }
    })
    return { point: best, clearance: Math.max(0, bestClearance), bounds }
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
    const sorted = values.filter(Number.isFinite).sort((a, b) => a - b)
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
      // Parallel, non-collinear segments have NO intersection. Returning a
      // false collinear flag without t/u poisons split sorting with undefined.
      return Math.abs(crossVector(qp, r)) <= epsilon ? { collinear: true } : null
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

  function cubicBezierDerivative(curve, t) {
    const mt = 1 - t
    return {
      x: 3 * mt * mt * (curve.p1.x - curve.p0.x) +
        6 * mt * t * (curve.p2.x - curve.p1.x) +
        3 * t * t * (curve.p3.x - curve.p2.x),
      y: 3 * mt * mt * (curve.p1.y - curve.p0.y) +
        6 * mt * t * (curve.p2.y - curve.p1.y) +
        3 * t * t * (curve.p3.y - curve.p2.y)
    }
  }

  function refineCurveLineIntersection(curveSegment, lineSegment, chordT) {
    const meta = curveSegment.sourceMeta
    if (!meta || meta.kind !== 'curve') return null
    const lineVector = subtractPoints(lineSegment.b, lineSegment.a)
    const curveParameter = localT => meta.t0 + (meta.t1 - meta.t0) * localT
    const signedDistance = localT => {
      const point = cubicBezierPoint(
        meta.curve.p0, meta.curve.p1, meta.curve.p2, meta.curve.p3, curveParameter(localT)
      )
      return crossVector(lineVector, subtractPoints(point, lineSegment.a))
    }
    let localT = Math.max(0, Math.min(1, chordT))
    for (let iteration = 0; iteration < 10; iteration += 1) {
      const globalT = curveParameter(localT)
      const value = signedDistance(localT)
      const derivative = cubicBezierDerivative(meta.curve, globalT)
      const slope = crossVector(lineVector, derivative) * (meta.t1 - meta.t0)
      if (Math.abs(slope) <= 1e-12) break
      const next = localT - value / slope
      if (!Number.isFinite(next) || next < -0.1 || next > 1.1) break
      localT = Math.max(0, Math.min(1, next))
    }
    const globalT = curveParameter(localT)
    const point = cubicBezierPoint(meta.curve.p0, meta.curve.p1, meta.curve.p2, meta.curve.p3, globalT)
    const lineT = segmentParameter(point, lineSegment)
    if (lineT < -1e-5 || lineT > 1 + 1e-5) return null
    return { curveT: localT, lineT: Math.max(0, Math.min(1, lineT)) }
  }

  function pointOnSourceSegment(segment, localT) {
    const meta = segment.sourceMeta
    if (!meta || meta.kind !== 'curve') return interpolatePoint(segment.a, segment.b, localT)
    const t = meta.t0 + (meta.t1 - meta.t0) * localT
    return cubicBezierPoint(meta.curve.p0, meta.curve.p1, meta.curve.p2, meta.curve.p3, t)
  }

  function refineCurveCurveIntersection(segmentA, segmentB, chordT, chordU) {
    const aMeta = segmentA.sourceMeta
    const bMeta = segmentB.sourceMeta
    let t = Math.max(0, Math.min(1, chordT))
    let u = Math.max(0, Math.min(1, chordU))
    for (let iteration = 0; iteration < 12; iteration += 1) {
      const pointA = pointOnSourceSegment(segmentA, t)
      const pointB = pointOnSourceSegment(segmentB, u)
      const delta = subtractPoints(pointA, pointB)
      const globalA = aMeta.t0 + (aMeta.t1 - aMeta.t0) * t
      const globalB = bMeta.t0 + (bMeta.t1 - bMeta.t0) * u
      const da0 = cubicBezierDerivative(aMeta.curve, globalA)
      const db0 = cubicBezierDerivative(bMeta.curve, globalB)
      const da = { x: da0.x * (aMeta.t1 - aMeta.t0), y: da0.y * (aMeta.t1 - aMeta.t0) }
      const db = { x: db0.x * (bMeta.t1 - bMeta.t0), y: db0.y * (bMeta.t1 - bMeta.t0) }
      const determinant = da.y * db.x - da.x * db.y
      if (Math.abs(determinant) <= 1e-14) break
      const dt = (delta.x * db.y - delta.y * db.x) / determinant
      const du = (da.y * delta.x - da.x * delta.y) / determinant
      if (!Number.isFinite(dt) || !Number.isFinite(du)) break
      t = Math.max(0, Math.min(1, t + dt))
      u = Math.max(0, Math.min(1, u + du))
      if (Math.abs(dt) + Math.abs(du) < 1e-10) break
    }
    const pa = pointOnSourceSegment(segmentA, t)
    const pb = pointOnSourceSegment(segmentB, u)
    return Math.hypot(pa.x - pb.x, pa.y - pb.y) <= 1e-4 ? { t, u } : null
  }


  function sourceSegmentBounds(segment) {
    const meta = segment.sourceMeta
    const points = meta && meta.kind === 'curve'
      ? Object.values(cubicSubcurve(meta.curve, meta.t0, meta.t1))
      : [segment.a, segment.b]
    return { left: Math.min(...points.map(p => p.x)), right: Math.max(...points.map(p => p.x)),
      top: Math.min(...points.map(p => p.y)), bottom: Math.max(...points.map(p => p.y)) }
  }

  // Sampling chords locate possible crossings, but are not the original curves.
  // A true crossing can lie in the neighbouring sample interval. Intersect the
  // actual subcurves instead, using their control hulls for conservative bounds.
  function sourceSegmentIntersections(a, b, epsilon) {
    const hits = []
    const curved = s => s.sourceMeta && s.sourceMeta.kind === 'curve'
    const flatness = s => {
      if (!curved(s)) return 0
      const c = cubicSubcurve(s.sourceMeta.curve, s.sourceMeta.t0, s.sourceMeta.t1)
      const d = subtractPoints(c.p3, c.p0)
      const length = Math.hypot(d.x, d.y)
      return length < epsilon ? Math.max(Math.hypot(c.p1.x-c.p0.x,c.p1.y-c.p0.y), Math.hypot(c.p2.x-c.p0.x,c.p2.y-c.p0.y))
        : Math.max(Math.abs(crossVector(d, subtractPoints(c.p1,c.p0))), Math.abs(crossVector(d, subtractPoints(c.p2,c.p0)))) / length
    }
    const half = s => {
      const mid = pointOnSourceSegment(s, 0.5)
      const meta = s.sourceMeta
      const tm = (meta.t0 + meta.t1) / 2
      return [{ a:s.a, b:mid, sourceMeta:{...meta,t1:tm} }, { a:mid, b:s.b, sourceMeta:{...meta,t0:tm} }]
    }
    const visit = (sa, sb, a0, a1, b0, b1, depth) => {
      const ba = sourceSegmentBounds(sa), bb = sourceSegmentBounds(sb)
      if (ba.right + epsilon < bb.left || bb.right + epsilon < ba.left ||
          ba.bottom + epsilon < bb.top || bb.bottom + epsilon < ba.top) return
      const fa = flatness(sa), fb = flatness(sb)
      if (depth < 24 && Math.max(fa, fb) > epsilon * 0.1) {
        if (fa >= fb) {
          const halves = half(sa), mid = (a0+a1)/2
          visit(halves[0],sb,a0,mid,b0,b1,depth+1); visit(halves[1],sb,mid,a1,b0,b1,depth+1)
        } else {
          const halves = half(sb), mid = (b0+b1)/2
          visit(sa,halves[0],a0,a1,b0,mid,depth+1); visit(sa,halves[1],a0,a1,mid,b1,depth+1)
        }
        return
      }
      const hit = segmentIntersection(sa, sb, 1e-12)
      if (!hit || hit.collinear) return
      let t = a0 + (a1-a0)*hit.t, u = b0 + (b1-b0)*hit.u
      if (curved(a) && curved(b)) {
        const refined = refineCurveCurveIntersection(a,b,t,u)
        if (!refined) return
        t=refined.t; u=refined.u
      } else {
        const refined = curved(a) ? refineCurveLineIntersection(a,b,t) : refineCurveLineIntersection(b,a,u)
        if (!refined) return
        t=curved(a)?refined.curveT:refined.lineT; u=curved(a)?refined.lineT:refined.curveT
      }
      const pa=pointOnSourceSegment(a,t), pb=pointOnSourceSegment(b,u)
      if (Math.hypot(pa.x-pb.x,pa.y-pb.y) <= epsilon &&
          !hits.some(h => Math.abs(h.t-t)<1e-8 && Math.abs(h.u-u)<1e-8)) hits.push({t,u})
    }
    visit(a,b,0,1,0,1,0)
    return hits
  }

  function genericPathsFromShape(shape, segments) {
    if (isPluginShape(shape)) return []
    const ooxmlPaths = ooxmlPathsForShape(shape, segments)
      .filter(path => path.stroke !== false && path.points && path.points.length >= (path.closed ? 3 : 2))
      .map(path => Object.assign({}, path, { shape, name: path.name || shapeName(shape) }))
    if (ooxmlPaths.length) return ooxmlPaths
    if (isCurvedConnector(shape)) return [sampleCurvedConnectorPath(shape, segments)]
    if (isStraightLine(shape)) {
      const line = lineFromShape(shape)
      return [{ shape, name: line.name, points: [line.p1, line.p2], closed: false, curved: false }]
    }
    if (getShapeType(shape) === M.msoFreeform) {
      return [sampleFreeformPath(shape, segments)]
    }
    if (supportedBoundaryKind(shape)) {
      const boundary = boundaryFromShape(shape, segments)
      return [{
        shape,
        name: boundary.name,
        points: boundary.points,
        closed: true,
        curved: boundary.kind === '椭圆' || boundary.kind === '圆角矩形'
      }]
    }
    return []
  }

  function genericPathFromShape(shape, segments) {
    const paths = genericPathsFromShape(shape, segments)
    return paths.length ? paths[0] : null
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
    if (!usablePaths.length) throw new Error('没有可读取的线条或闭合形状')

    const allPoints = usablePaths.reduce((result, path) => result.concat(path.points), [])
    const xs = allPoints.map(point => point.x)
    const ys = allPoints.map(point => point.y)
    const scale = Math.max(
      Math.max.apply(null, xs) - Math.min.apply(null, xs),
      Math.max.apply(null, ys) - Math.min.apply(null, ys),
      1
    )
    const epsilon = scale * 1e-7
    // Distinct nearby intersections must remain distinct. A visual-size snap
    // changes the cyclic edge order at a junction (e.g. can + bevel), producing
    // face walks that retrace curves and join unrelated regions. Only deduplicate
    // numerical copies here; marker/display tolerances must not alter topology.
    const mergeTolerance = Math.max(1e-6, epsilon * 4)
    const segments = []
    usablePaths.forEach((path, pathIndex) => {
      const edgeCount = path.closed ? path.points.length : path.points.length - 1
      for (let i = 0; i < edgeCount; i += 1) {
        const a = path.points[i]
        const b = path.points[(i + 1) % path.points.length]
        if (Math.hypot(b.x - a.x, b.y - a.y) <= epsilon) continue
        segments.push({
          a,
          b,
          pathIndex,
          sourceMeta: path.edgeMeta && path.edgeMeta[i] ? path.edgeMeta[i] : null,
          splits: [0, 1]
        })
      }
    })
    if (segments.length > 6000) {
      throw new Error('线稿采样后超过 6000 段，请降低曲边精度或减少参与识别的对象。')
    }

    const bounds = segments.map(segment => {
      const box = sourceSegmentBounds(segment)
      return { left:box.left-epsilon, right:box.right+epsilon, top:box.top-epsilon, bottom:box.bottom+epsilon }
    })
    for (let i = 0; i < segments.length; i += 1) {
      for (let j = i + 1; j < segments.length; j += 1) {
        if (
          bounds[i].right < bounds[j].left || bounds[j].right < bounds[i].left ||
          bounds[i].bottom < bounds[j].top || bounds[j].bottom < bounds[i].top
        ) continue
        const iCurve = segments[i].sourceMeta && segments[i].sourceMeta.kind === 'curve'
        const jCurve = segments[j].sourceMeta && segments[j].sourceMeta.kind === 'curve'
        if (iCurve || jCurve) {
          sourceSegmentIntersections(segments[i], segments[j], epsilon).forEach(hit => {
            segments[i].splits.push(hit.t)
            segments[j].splits.push(hit.u)
          })
          continue
        }
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
    const bucketSize = mergeTolerance
    const bucketKey = (x, y) => x + ':' + y
    const getVertexId = point => {
      const bx = Math.round(point.x / bucketSize)
      const by = Math.round(point.y / bucketSize)
      for (let dx = -1; dx <= 1; dx += 1) {
        for (let dy = -1; dy <= 1; dy += 1) {
          const ids = vertexBuckets.get(bucketKey(bx + dx, by + dy)) || []
          for (const id of ids) {
            if (Math.hypot(vertices[id].x - point.x, vertices[id].y - point.y) <= mergeTolerance) return id
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
        const a = pointOnSourceSegment(segment, splits[i])
        const b = pointOnSourceSegment(segment, splits[i + 1])
        if (Math.hypot(b.x - a.x, b.y - a.y) <= epsilon) continue
        const u = getVertexId(a)
        const v = getVertexId(b)
        if (u === v) continue
        const key = u < v ? u + ':' + v : v + ':' + u
        if (!edges.has(key)) {
          let sourceMeta = null
          if (segment.sourceMeta && segment.sourceMeta.kind === 'curve') {
            sourceMeta = {
              kind: 'curve',
              curve: segment.sourceMeta.curve,
              t0: segment.sourceMeta.t0 + (segment.sourceMeta.t1 - segment.sourceMeta.t0) * splits[i],
              t1: segment.sourceMeta.t0 + (segment.sourceMeta.t1 - segment.sourceMeta.t0) * splits[i + 1]
            }
          }
          edges.set(key, { u, v, sourceMeta })
        }
      }
    })

    const graphEdges = Array.from(edges.values())
    const components = vertices.map((_, index) => index)
    const root = id => {
      while (components[id] !== id) { components[id] = components[components[id]]; id = components[id] }
      return id
    }
    graphEdges.forEach(edge => { components[root(edge.u)] = root(edge.v) })
    const edgeByKey = new Map()
    graphEdges.forEach(edge => {
      edgeByKey.set(edge.u < edge.v ? edge.u + ':' + edge.v : edge.v + ':' + edge.u, edge)
    })
    const regions = enumerateGraphFaces(vertices, graphEdges, true)
      .map(record => {
        const region = record.points.slice()
        region._component = root(record.nodeIds[0])
        region._pathEdges = record.nodeIds.map((u, index) => {
          const v = record.nodeIds[(index + 1) % record.nodeIds.length]
          const edge = edgeByKey.get(u < v ? u + ':' + v : v + ':' + u)
          if (!edge || !edge.sourceMeta) return { kind: 'line' }
          const forward = edge.u === u && edge.v === v
          return {
            kind: 'curve',
            curve: edge.sourceMeta.curve,
            t0: forward ? edge.sourceMeta.t0 : edge.sourceMeta.t1,
            t1: forward ? edge.sourceMeta.t1 : edge.sourceMeta.t0
          }
        })
        return region
      })
      .filter(region => region.length >= 3 && Math.abs(polygonArea(region)) > scale * scale * 1e-8)
    if (!regions.length) {
      throw new Error('通用线稿中没有形成闭合区域；请确认线条端点相接或彼此相交。')
    }
    // Disconnected nested outlines delimit holes, not overlapping filled disks.
    // Keep each inner face selectable, but subtract it from its nearest parent.
    regions.forEach(child => {
      const childArea = Math.abs(polygonArea(child))
      const parents = regions.filter(parent => parent._component !== child._component &&
        Math.abs(polygonArea(parent)) > childArea + epsilon &&
        child.every(point => pointInPolygon(point, parent, epsilon)))
        .sort((a, b) => Math.abs(polygonArea(a)) - Math.abs(polygonArea(b)))
      if (parents.length) {
        const parent = parents[0]
        if (!parent._holes) parent._holes = []
        parent._holes.push(child)
      }
    })
    Object.defineProperty(regions, '_debug', {
      value: { vertices, graphEdges, mergeTolerance },
      enumerable: false,
      configurable: true
    })
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
    const native = { start: null, commands: [] }
    const arcBuilder = global.WpsOoxmlGeometry && global.WpsOoxmlGeometry._internals.arcCubics
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
      if (arcBuilder) {
        const start = ellipsePointOnRay(current, startAngle)
        if (!native.start) native.start = start
        else native.commands.push({ kind: 'line', end: start })
        arcBuilder(start, current.rx, current.ry, startAngle * 180 / Math.PI * 60000, -clockwiseGap * 180 / Math.PI * 60000)
          .forEach(curve => native.commands.push({ kind: 'curve', p1: curve.p1, p2: curve.p2, end: curve.p3 }))
      }
      const arcSteps = Math.max(2, Math.ceil(segments * clockwiseGap / fullTurn))
      for (let step = 0; step <= arcSteps; step += 1) {
        const angle = startAngle - clockwiseGap * step / arcSteps
        pushDistinct(points, ellipsePointOnRay(current, angle), 1e-6)
      }
    }
    const polygon = cleanPolygon(points, 1e-6)
    if (native.start) {
      native.commands.push({ kind: 'line', end: native.start })
      polygon._nativePath = native
    }
    return polygon
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
      ? nodes.map(node => {
        const polygon = sampleEllipse(node, segments)
        const source = ooxmlPathsForShape(node.shape, segments).find(path => path.closed)
        if (source) {
          const exact = source.points.slice()
          exact._pathEdges = source.edgeMeta
          return exact
        }
        return polygon
      })
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
    if (boundaryOrEllipse && boundaryOrEllipse.sourcePath) {
      const path = boundaryOrEllipse.sourcePath
      return buildGenericRegions([path].concat((lines || []).map(line => ({
        points: [line.p1, line.p2], closed: false
      })))).filter(region => pointInPolygon(polygonInteriorPoint(region).point, path.points))
    }
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

  function setMarkerStyle(shape, picked, colorHex) {
    shape.Fill.Visible = M.msoTrue
    try { shape.Fill.Solid() } catch (_) {}
    shape.Fill.ForeColor.RGB = officeRgb(picked ? colorHex : '#ffffff')
    shape.Fill.Transparency = 0
    shape.Line.Visible = M.msoTrue
    try { shape.Line.ForeColor.RGB = officeRgb(picked ? '#ffffff' : colorHex) } catch (_) {}
    try { shape.Line.Weight = picked ? 2.25 : 1.75 } catch (_) {}
  }

  function createSelectionMarker(slide, polygon, key, colorHex) {
    const interior = polygonInteriorPoint(polygon)
    const minDimension = Math.min(
      interior.bounds.right - interior.bounds.left,
      interior.bounds.bottom - interior.bounds.top
    )
    const diameter = Math.max(6, Math.min(14, interior.clearance * 0.9, minDimension * 0.32))
    const left = interior.point.x - diameter / 2
    const top = interior.point.y - diameter / 2
    let marker = null
    try {
      if (slide.Shapes && typeof slide.Shapes.AddShape === 'function') {
        marker = slide.Shapes.AddShape(M.msoShapeOval, left, top, diameter, diameter)
      }
    } catch (_) {
      marker = null
    }
    if (!marker) {
      const radius = diameter / 2
      marker = createFreeform(slide, [
        { x: interior.point.x, y: interior.point.y - radius },
        { x: interior.point.x + radius, y: interior.point.y },
        { x: interior.point.x, y: interior.point.y + radius },
        { x: interior.point.x - radius, y: interior.point.y }
      ], '#ffffff', 100, MARKER_PREFIX + key)
    }
    marker.Name = MARKER_PREFIX + key
    setMarkerStyle(marker, false, colorHex)
    try { marker.ZOrder(M.msoBringToFront) } catch (_) {}
    return marker
  }

  function splitCubic(curve, t) {
    const p01 = interpolatePoint(curve.p0, curve.p1, t)
    const p12 = interpolatePoint(curve.p1, curve.p2, t)
    const p23 = interpolatePoint(curve.p2, curve.p3, t)
    const p012 = interpolatePoint(p01, p12, t)
    const p123 = interpolatePoint(p12, p23, t)
    const p = interpolatePoint(p012, p123, t)
    return {
      left: { p0: curve.p0, p1: p01, p2: p012, p3: p },
      right: { p0: p, p1: p123, p2: p23, p3: curve.p3 }
    }
  }

  function cubicSubcurve(curve, t0, t1) {
    const reverse = t1 < t0
    const start = Math.max(0, Math.min(1, reverse ? t1 : t0))
    const end = Math.max(0, Math.min(1, reverse ? t0 : t1))
    let result
    if (end <= 1e-12) {
      result = { p0: curve.p0, p1: curve.p0, p2: curve.p0, p3: curve.p0 }
    } else {
      const prefix = splitCubic(curve, end).left
      result = start <= 1e-12 ? prefix : splitCubic(prefix, start / end).right
    }
    if (!reverse) return result
    return { p0: result.p3, p1: result.p2, p2: result.p1, p3: result.p0 }
  }

  function freeformCommands(polygon) {
    const edges = polygon && polygon._pathEdges
    if (!edges || edges.length !== polygon.length || !edges.some(edge => edge.kind === 'curve')) return null
    let startIndex = edges.findIndex(edge => edge.kind !== 'curve')
    if (startIndex < 0) startIndex = 0
    const commands = []
    for (let offset = 0; offset < edges.length; offset += 1) {
      const index = (startIndex + offset) % edges.length
      const edge = edges[index]
      const endPoint = polygon[(index + 1) % polygon.length]
      const previous = commands[commands.length - 1]
      const sameDirection = previous && previous.kind === 'curve' && edge.kind === 'curve' &&
        Math.sign(previous.t1 - previous.t0) === Math.sign(edge.t1 - edge.t0)
      if (
        sameDirection && previous.curve === edge.curve &&
        Math.abs(previous.t1 - edge.t0) <= 1e-6
      ) {
        previous.t1 = edge.t1
        previous.end = endPoint
      } else {
        commands.push(edge.kind === 'curve'
          ? { kind: 'curve', curve: edge.curve, t0: edge.t0, t1: edge.t1, end: endPoint }
          : { kind: 'line', end: endPoint })
      }
    }
    return { start: polygon[startIndex], commands }
  }

  function buildCurvedFreeform(slide, polygon, offsetX, offsetY) {
    const path = freeformCommands(polygon)
    if (!path) return null
    const dx = toNumber(offsetX, 0)
    const dy = toNumber(offsetY, 0)
    const builder = slide.Shapes.BuildFreeform(M.msoEditingCorner, path.start.x - dx, path.start.y - dy)
    path.commands.forEach(command => {
      if (command.kind === 'curve') {
        const curve = cubicSubcurve(command.curve, command.t0, command.t1)
        builder.AddNodes(
          M.msoSegmentCurve,
          M.msoEditingCorner,
          curve.p1.x - dx,
          curve.p1.y - dy,
          curve.p2.x - dx,
          curve.p2.y - dy,
          curve.p3.x - dx,
          curve.p3.y - dy
        )
      } else {
        builder.AddNodes(
          M.msoSegmentLine,
          M.msoEditingAuto,
          command.end.x - dx,
          command.end.y - dy
        )
      }
    })
    return builder.ConvertToShape()
  }

  function addCurvePointRows(polygon) {
    const path = freeformCommands(polygon)
    if (!path) return null
    const rows = [[path.start.x, path.start.y]]
    let current = path.start
    path.commands.forEach(command => {
      if (command.kind === 'curve') {
        const curve = cubicSubcurve(command.curve, command.t0, command.t1)
        rows.push(
          [curve.p1.x, curve.p1.y],
          [curve.p2.x, curve.p2.y],
          [curve.p3.x, curve.p3.y]
        )
        current = curve.p3
      } else {
        // Shapes.AddCurve 只接受三次贝塞尔点列。直线可以用位于 1/3、2/3
        // 位置的两个共线控制点精确表示，不会产生任何几何近似。
        const end = command.end
        rows.push(
          [current.x + (end.x - current.x) / 3, current.y + (end.y - current.y) / 3],
          [current.x + (end.x - current.x) * 2 / 3, current.y + (end.y - current.y) * 2 / 3],
          [end.x, end.y]
        )
        current = end
      }
    })
    return rows
  }

  function buildCurvedShape(slide, polygon, expected) {
    const failures = []
    const rows = addCurvePointRows(polygon)

    // WPS 12.1 的 FreeformBuilder.AddNodes 在部分安装中会丢失曲线调用的
    // X3/Y3 参数。AddCurve 一次传入完整点列，不经过该可选参数桥接，优先使用。
    if (rows && slide.Shapes && typeof slide.Shapes.AddCurve === 'function') {
      let shape = null
      try {
        shape = slide.Shapes.AddCurve(rows)
        if (shape && freeformBoundsMatch(shape, expected)) return shape
        failures.push('AddCurve 返回的范围与目标区域不一致')
      } catch (error) {
        failures.push('AddCurve：' + String(error && error.message ? error.message : error))
      }
      discardShape(shape)
    } else {
      failures.push('当前 WPS 未提供 Shapes.AddCurve')
    }

    let shape = null
    try {
      shape = buildCurvedFreeform(slide, polygon, 0, 0)
      if (shape && freeformBoundsMatch(shape, expected)) return shape
      failures.push('BuildFreeform 返回的范围与目标区域不一致')
    } catch (error) {
      failures.push('BuildFreeform：' + String(error && error.message ? error.message : error))
    }
    discardShape(shape)
    const detail = failures.filter(Boolean).join('；')
    throw new Error(
      'WPS 未能按 PPTX 中的真实贝塞尔控制点创建曲边填色块；已停止生成，未使用折线替代。' +
      (detail ? ' 技术信息：' + detail : '')
    )
  }

  function freeformBounds(shape, expected) {
    const left = toNumber(shape && shape.Left, NaN)
    const top = toNumber(shape && shape.Top, NaN)
    const width = Math.abs(toNumber(shape && shape.Width, NaN))
    const height = Math.abs(toNumber(shape && shape.Height, NaN))
    if (![left, top, width, height].every(Number.isFinite)) return null
    return { left, top, right: left + width, bottom: top + height, width, height }
  }

  function freeformBoundsMatch(shape, expected) {
    const actual = freeformBounds(shape, expected)
    if (!actual) return false
    const expectedWidth = Math.max(expected.right - expected.left, 1e-6)
    const expectedHeight = Math.max(expected.bottom - expected.top, 1e-6)
    const expectedSize = Math.max(expectedWidth, expectedHeight, 1)
    const tolerance = Math.max(8, expectedSize * 0.08)
    return !(
      actual.left < expected.left - tolerance ||
      actual.top < expected.top - tolerance ||
      actual.right > expected.right + tolerance ||
      actual.bottom > expected.bottom + tolerance ||
      actual.width < expectedWidth * 0.45 ||
      actual.height < expectedHeight * 0.45
    )
  }

  function discardShape(shape) {
    if (!shape) return
    try { shape.Delete() } catch (_) {}
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

    // 真实曲边区域必须使用原始 OOXML 贝塞尔控制点生成，不能再把采样点
    // 当作大量直线节点。WPS 的 Nodes.Item() 读取接口本身会返回 null，
    // 因此这里只用生成后的空间范围校验结果，不再用 Nodes 反向判定曲线。
    if (freeformCommands(polygon)) {
      shape = buildCurvedShape(slide, polygon, expected)
    }

    if (!shape) {
      try {
        if (slide.Shapes && typeof slide.Shapes.AddPolyline === 'function') {
          shape = slide.Shapes.AddPolyline(closedPoints.map(point => [point.x, point.y]))
        }
      } catch (_) {
        shape = null
      }
    }

    if (!shape) {
      const builder = slide.Shapes.BuildFreeform(M.msoEditingCorner, first.x, first.y)
      for (let i = 1; i < polygon.length; i += 1) {
        const point = polygon[i]
        builder.AddNodes(
          M.msoSegmentLine,
          M.msoEditingCorner,
          point.x, point.y,
          point.x, point.y,
          point.x, point.y
        )
      }
      builder.AddNodes(
        M.msoSegmentLine,
        M.msoEditingCorner,
        first.x, first.y,
        first.x, first.y,
        first.x, first.y
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
      discardShape(shape)
      throw new Error('WPS 返回的自由形状坐标超出目标区域，已停止生成以避免产生错位对象。')
    }
    shape.Name = name
    setFill(shape, colorHex, opacityPercent)
    shape.ZOrder(M.msoSendToBack)
    return shape
  }

  function clearGeometryDebugShapes(slide) {
    let deleted = 0
    for (let i = toNumber(slide && slide.Shapes && slide.Shapes.Count, 0); i >= 1; i -= 1) {
      const shape = slide.Shapes.Item(i)
      if (shapeName(shape).indexOf(DEBUG_PREFIX) === 0) {
        shape.Delete()
        deleted += 1
      }
    }
    return deleted
  }

  function debugArrangementPieces(debug) {
    const vertices = debug.vertices || []
    const edges = debug.graphEdges || []
    const degree = Array(vertices.length).fill(0)
    edges.forEach(edge => {
      degree[edge.u] += 1
      degree[edge.v] += 1
    })
    const linePieces = edges.filter(edge => !edge.sourceMeta || edge.sourceMeta.kind !== 'curve')
    const curveGroups = new Map()
    edges.forEach(edge => {
      if (!edge.sourceMeta || edge.sourceMeta.kind !== 'curve') return
      const curve = edge.sourceMeta.curve
      if (!curveGroups.has(curve)) curveGroups.set(curve, [])
      curveGroups.get(curve).push(edge)
    })
    const curvePieces = []
    curveGroups.forEach(group => {
      group.sort((a, b) => a.sourceMeta.t0 - b.sourceMeta.t0)
      group.forEach(edge => {
        const previous = curvePieces[curvePieces.length - 1]
        if (
          previous && previous.sourceMeta.curve === edge.sourceMeta.curve &&
          previous.v === edge.u && degree[edge.u] === 2 &&
          Math.abs(previous.sourceMeta.t1 - edge.sourceMeta.t0) <= 1e-6
        ) {
          previous.v = edge.v
          previous.sourceMeta.t1 = edge.sourceMeta.t1
        } else {
          curvePieces.push({
            u: edge.u,
            v: edge.v,
            sourceMeta: {
              kind: 'curve',
              curve: edge.sourceMeta.curve,
              t0: edge.sourceMeta.t0,
              t1: edge.sourceMeta.t1
            }
          })
        }
      })
    })
    const intersections = vertices
      .map((point, index) => ({ point, index, degree: degree[index] }))
      .filter(item => item.degree > 2)
    return { vertices, degree, linePieces, curvePieces, intersections }
  }

  function styleDebugSegment(shape, colorHex, weight) {
    try { shape.Fill.Visible = M.msoFalse } catch (_) {}
    try { shape.Line.Visible = M.msoTrue } catch (_) {}
    try { shape.Line.ForeColor.RGB = officeRgb(colorHex) } catch (_) {}
    try { shape.Line.Weight = weight } catch (_) {}
    try { shape.ZOrder(M.msoBringToFront) } catch (_) {}
  }

  function createDebugLineSegment(slide, a, b, name) {
    let shape = null
    if (slide.Shapes && typeof slide.Shapes.AddLine === 'function') {
      shape = slide.Shapes.AddLine(a.x, a.y, b.x, b.y)
    } else {
      const builder = slide.Shapes.BuildFreeform(M.msoEditingCorner, a.x, a.y)
      builder.AddNodes(M.msoSegmentLine, M.msoEditingAuto, b.x, b.y)
      shape = builder.ConvertToShape()
    }
    shape.Name = name
    styleDebugSegment(shape, '#1769aa', 2.25)
    return shape
  }

  function createDebugCurveSegment(slide, a, edge, name) {
    const meta = edge.sourceMeta
    const span = Math.abs(meta.t1 - meta.t0)
    const steps = Math.max(6, Math.min(240, Math.ceil(span * 240)))
    const points = []
    for (let i = 0; i <= steps; i += 1) {
      const t = meta.t0 + (meta.t1 - meta.t0) * i / steps
      points.push(cubicBezierPoint(
        meta.curve.p0,
        meta.curve.p1,
        meta.curve.p2,
        meta.curve.p3,
        t
      ))
    }
    // 调试对象恢复使用原曲线参数区间的稳定采样；不把控制点再次提交给
    // 当前 WPS 版本不稳定的 BuildFreeform 曲线节点接口。
    const shape = slide.Shapes.AddPolyline(points.map(point => [point.x, point.y]))
    shape.Name = name
    styleDebugSegment(shape, '#ef6c00', 3)
    return { shape, curveStatus: null }
  }

  function createDebugIntersection(slide, point, index) {
    const diameter = 8
    const marker = slide.Shapes.AddShape(
      M.msoShapeOval,
      point.x - diameter / 2,
      point.y - diameter / 2,
      diameter,
      diameter
    )
    marker.Name = DEBUG_PREFIX + 'Point_' + String(index + 1).padStart(2, '0')
    marker.Fill.Visible = M.msoTrue
    try { marker.Fill.Solid() } catch (_) {}
    marker.Fill.ForeColor.RGB = officeRgb('#e53935')
    marker.Fill.Transparency = 0
    marker.Line.Visible = M.msoTrue
    try { marker.Line.ForeColor.RGB = officeRgb('#ffffff') } catch (_) {}
    try { marker.Line.Weight = 1.5 } catch (_) {}
    try { marker.ZOrder(M.msoBringToFront) } catch (_) {}
    return marker
  }

  function diagnosticNumbers(value) {
    const numbers = []
    try { flattenNumbers(value, numbers, 0) } catch (_) {}
    return numbers.map(value => Math.round(value * 1000000) / 1000000)
  }

  function diagnosticProperty(object, name) {
    try {
      const value = object && object[name]
      if (value === null || value === undefined) return null
      if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value
      return String(value)
    } catch (error) {
      return { error: String(error && error.message ? error.message : error) }
    }
  }

  function inspectShapeGeometry(shape) {
    const report = {
      name: shapeName(shape),
      id: diagnosticProperty(shape, 'Id'),
      type: diagnosticProperty(shape, 'Type'),
      autoShapeType: diagnosticProperty(shape, 'AutoShapeType'),
      connector: diagnosticProperty(shape, 'Connector'),
      left: diagnosticProperty(shape, 'Left'),
      top: diagnosticProperty(shape, 'Top'),
      width: diagnosticProperty(shape, 'Width'),
      height: diagnosticProperty(shape, 'Height'),
      rotation: diagnosticProperty(shape, 'Rotation'),
      horizontalFlip: diagnosticProperty(shape, 'HorizontalFlip'),
      verticalFlip: diagnosticProperty(shape, 'VerticalFlip')
    }
    try {
      report.connectorFormat = {
        type: diagnosticProperty(shape.ConnectorFormat, 'Type'),
        beginConnected: diagnosticProperty(shape.ConnectorFormat, 'BeginConnected'),
        endConnected: diagnosticProperty(shape.ConnectorFormat, 'EndConnected'),
        beginConnectionSite: diagnosticProperty(shape.ConnectorFormat, 'BeginConnectionSite'),
        endConnectionSite: diagnosticProperty(shape.ConnectorFormat, 'EndConnectionSite')
      }
    } catch (error) {
      report.connectorFormat = { error: String(error && error.message ? error.message : error) }
    }
    try {
      report.vertices = diagnosticNumbers(shape.Vertices)
    } catch (error) {
      report.vertices = { error: String(error && error.message ? error.message : error) }
    }
    try {
      const nodes = shape.Nodes
      const count = toNumber(nodes && nodes.Count, 0)
      report.nodes = []
      for (let i = 1; i <= count; i += 1) {
        const node = nodes.Item(i)
        if (!node) {
          report.nodes.push({ index: i, error: 'Item returned null' })
          continue
        }
        report.nodes.push({
          index: i,
          segmentType: diagnosticProperty(node, 'SegmentType'),
          editingType: diagnosticProperty(node, 'EditingType'),
          points: diagnosticNumbers(node.Points)
        })
      }
      if (!count) report.nodesStatus = 'Count=' + String(diagnosticProperty(nodes, 'Count'))
    } catch (error) {
      report.nodes = { error: String(error && error.message ? error.message : error) }
    }
    try {
      const adjustments = shape.Adjustments
      const count = toNumber(adjustments && adjustments.Count, 0)
      report.adjustments = []
      for (let i = 1; i <= count; i += 1) {
        try { report.adjustments.push(Number(adjustments.Item(i))) } catch (error) {
          report.adjustments.push({ error: String(error && error.message ? error.message : error) })
        }
      }
    } catch (error) {
      report.adjustments = { error: String(error && error.message ? error.message : error) }
    }
    try {
      const id = String(shape.Id)
      const entry = geometrySnapshot && geometrySnapshot.shapes && geometrySnapshot.shapes[id]
      const paths = ooxmlPathsForShape(shape, 64)
      report.pptxGeometry = entry
        ? {
            source: 'pptx-ooxml',
            preset: entry.preset || null,
            adjustments: entry.adjustments || {},
            pathCount: paths.length,
            closedPathCount: paths.filter(path => path.closed).length,
            curvedPathCount: paths.filter(path => path.curved).length,
            pointCount: paths.reduce((sum, path) => sum + path.points.length, 0)
          }
        : { source: 'unavailable' }
    } catch (error) {
      report.pptxGeometry = { error: String(error && error.message ? error.message : error) }
    }
    if (isCurvedConnector(shape)) {
      try {
        const sampled = sampleCurvedConnectorPath(shape, 32)
        report.pluginCurveInterpretation = {
          source: sampled.curveSource,
          pointCount: sampled.points.length,
          first: sampled.points[0],
          middle: sampled.points[Math.floor(sampled.points.length / 2)],
          last: sampled.points[sampled.points.length - 1]
        }
      } catch (error) {
        report.pluginCurveInterpretation = { error: String(error && error.message ? error.message : error) }
      }
    }
    return report
  }

  function inspectSelectedGeometry() {
    const app = getApplication()
    const slide = getActiveSlide(app)
    const selected = expandShapes(getSelectedShapes(
      app,
      '请先选中需要诊断的曲线连接符和相关轮廓。'
    )).filter(shape => !isPluginShape(shape))
    if (!selected.length) throw new Error('请先选中至少一个原始对象，不要选中插件生成的调试对象。')
    const report = {
      application: {
        name: diagnosticProperty(app, 'Name'),
        version: diagnosticProperty(app, 'Version'),
        build: diagnosticProperty(app, 'Build')
      },
      curveCreationApi: {
        addCurve: Boolean(slide && slide.Shapes && typeof slide.Shapes.AddCurve === 'function'),
        buildFreeform: Boolean(slide && slide.Shapes && typeof slide.Shapes.BuildFreeform === 'function'),
        addPolyline: Boolean(slide && slide.Shapes && typeof slide.Shapes.AddPolyline === 'function')
      },
      generatedAt: new Date().toISOString(),
      shapes: selected.map(inspectShapeGeometry)
    }
    return JSON.stringify(report, null, 2)
  }

  function generateGeometryDebug(options) {
    const model = collectSelectionModel(options || {})
    if (model.mode !== 'generic' || !model.regions || !model.regions._debug) {
      throw new Error('几何调试当前只用于通用线稿模式；请同时选中需要检查的闭合轮廓、直线和曲线。')
    }
    clearGeometryDebugShapes(model.slide)
    const pieces = debugArrangementPieces(model.regions._debug)
    let curveConfirmed = 0
    let curveUnknown = 0
    let curveRejected = 0
    pieces.linePieces.forEach((edge, index) => {
      createDebugLineSegment(
        model.slide,
        pieces.vertices[edge.u],
        pieces.vertices[edge.v],
        DEBUG_PREFIX + 'Line_' + String(index + 1).padStart(3, '0')
      )
    })
    pieces.curvePieces.forEach((edge, index) => {
      try {
        const result = createDebugCurveSegment(
          model.slide,
          pieces.vertices[edge.u],
          edge,
          DEBUG_PREFIX + 'Curve_' + String(index + 1).padStart(3, '0')
        )
        if (result.curveStatus === true) curveConfirmed += 1
        else if (result.curveStatus === false) curveRejected += 1
        else curveUnknown += 1
      } catch (_) {
        curveRejected += 1
      }
    })
    pieces.intersections.forEach((item, index) => createDebugIntersection(model.slide, item.point, index))
    unselectCurrentShapes(model.app)
    return {
      intersectionCount: pieces.intersections.length,
      lineSegmentCount: pieces.linePieces.length,
      curveSegmentCount: pieces.curvePieces.length,
      curveConfirmed,
      curveUnknown,
      curveRejected
    }
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

    if (ellipseShapes.length >= 3 && lineShapes.length >= 3 && boundaryShapes.length !== 1 &&
        ellipseShapes.every(shape => Math.abs(toNumber(shape.Rotation, 0) % 180) < 1e-6)) {
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

    const boundaryOnlyShape = boundaryShapes.length === 1 ? boundaryShapes[0] : null
    const boundaryModeCompatible = boundaryOnlyShape && sourceShapes.every(shape => (
      shape === boundaryOnlyShape || isStraightLine(shape)
    ))
    if (boundaryModeCompatible) {
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
    const paths = []
    sourceShapes.forEach(shape => {
      try {
        const shapePaths = genericPathsFromShape(shape, segments)
        if (!shapePaths.length) pathFailures.push(shapeName(shape) + '（PPTX 中没有可用路径）')
        shapePaths.forEach(path => paths.push(path))
      } catch (error) {
        pathFailures.push(shapeName(shape) + '（' + (error && error.message ? error.message : '读取失败') + '）')
      }
    })
    let genericRegions
    try {
      genericRegions = buildGenericRegions(paths)
    } catch (error) {
      const reason = String(error && error.message ? error.message : error)
        .replace(/[\s。；;，,]+$/g, '')
      const failed = pathFailures.length
        ? '；未读取对象：' + pathFailures.slice(0, 3).join('、')
        : ''
      throw new Error(
        '通用线稿扫描了 ' + sourceShapes.length + ' 个对象：成功读取 ' + paths.length +
        ' 条路径；' + reason + failed + '。'
      )
    }
    const closedShapes = []
    paths.filter(path => path.closed).forEach(path => {
      if (path.shape && closedShapes.indexOf(path.shape) < 0) closedShapes.push(path.shape)
    })
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
      if (isPreviewShape(shape) || isMarkerShape(shape)) {
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

  function markerSelectionKey(shape) {
    return shapeName(shape)
      .replace(PICKED_MARKER_PREFIX, '')
      .replace(MARKER_PREFIX, '')
  }

  function findPreviewByKey(slide, key) {
    for (let i = 1; i <= toNumber(slide.Shapes.Count, 0); i += 1) {
      const shape = slide.Shapes.Item(i)
      if (isPreviewShape(shape) && previewSelectionKey(shape) === key) return shape
    }
    return null
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

  function setMarkerPicked(shape, picked, color) {
    const key = markerSelectionKey(shape)
    shape.Name = (picked ? PICKED_MARKER_PREFIX : MARKER_PREFIX) + key
    setMarkerStyle(shape, picked, color)
  }

  function captureManualSelection(options) {
    const settings = Object.assign({ color: '#4f7cff' }, options || {})
    const app = getApplication()
    const slide = getActiveSlide(app)
    const selectedShapes = getSelectedShapes(app, '', true)
    const selected = selectedShapes.filter(isMarkerShape)

    if (!selected.length) {
      lastPreviewSelectionKey = ''
      if (selectedShapes.some(isPreviewShape)) {
        unselectCurrentShapes(app)
        return {
          changed: false,
          count: countPickedPreviews(slide),
          hint: '请点击候选区域内部的圆形定位点，不要点击整块浅色区域。'
        }
      }
      return { changed: false, count: countPickedPreviews(slide), hint: '' }
    }

    const selectionKey = selected.map(markerSelectionKey).sort().join('|')
    if (selectionKey === lastPreviewSelectionKey) {
      unselectCurrentShapes(app)
      return { changed: false, count: countPickedPreviews(slide) }
    }

    lastPreviewSelectionKey = selectionKey
    selected.forEach(marker => {
      const key = markerSelectionKey(marker)
      const preview = findPreviewByKey(slide, key)
      if (!preview) return
      const picked = !isPickedPreviewShape(preview)
      setPreviewPicked(preview, picked, settings.color)
      setMarkerPicked(marker, picked, settings.color)
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

  function nativePath(polygon) {
    if (polygon._nativePath) return polygon._nativePath
    const path = freeformCommands(polygon)
    if (!path) return { start: polygon[0], commands: polygon.slice(1).concat([polygon[0]]).map(end => ({ kind: 'line', end })) }
    return { start: path.start, commands: path.commands.map(command => {
      if (command.kind !== 'curve') return { kind: 'line', end: command.end }
      const curve = cubicSubcurve(command.curve, command.t0, command.t1)
      return { kind: 'curve', p1: curve.p1, p2: curve.p2, end: curve.p3 }
    }) }
  }

  function nativePaths(polygon) {
    const paths = [nativePath(polygon)]
    ;(polygon._holes || []).forEach(hole => {
      const path = nativePath(hole)
      const starts = [path.start].concat(path.commands.map(c => c.end))
      paths.push({ start: starts[starts.length - 1], commands: path.commands.map((c, i) => c.kind === 'curve'
        ? { kind: 'curve', p1: c.p2, p2: c.p1, end: starts[i] }
        : { kind: 'line', end: starts[i] }).reverse() })
    })
    return paths
  }

  async function generateNative(options, preview) {
    const settings = Object.assign({ color: '#4f7cff', opacity: 72, segments: 160, palette: false, disableBoundaryFill: true }, options || {})
    const model = collectSelectionModel(settings)
    if (!geometrySnapshot || !geometrySnapshot.createNativePackage || !global.WpsNativeShapeWriter) {
      throw new Error('请重新打开填色面板以加载原生曲边生成组件。')
    }
    const runId = Date.now().toString(36)
    const items = model.regions.map((polygon, index) => ({
      name: (preview ? PREVIEW_PREFIX : PREFIX) + runId + '_' + String(index + 1).padStart(2, '0'),
      paths: nativePaths(polygon),
      color: !preview && settings.palette ? paletteColor(settings.color, index, model.regions.length) : settings.color,
      opacity: preview ? 18 : settings.opacity
    }))
    const presentation = model.app.ActivePresentation
    const bytes = await geometrySnapshot.createNativePackage(items, Number(presentation.PageSetup.SlideWidth), Number(presentation.PageSetup.SlideHeight))
    const created = []
    try {
      created.push(...await global.WpsNativeShapeWriter.importPackage(model.app, model.slide, bytes, items))
      // Remove only older candidates after the new region batch was imported.
      for (let i = Number(model.slide.Shapes.Count); i >= 1; i--) {
        const shape = model.slide.Shapes.Item(i)
        if ((isPreviewShape(shape) || isMarkerShape(shape)) && !items.some(item => item.name === shapeName(shape))) shape.Delete()
      }
      if (preview) model.regions.forEach((polygon, index) => {
        created.push(createSelectionMarker(model.slide, polygon, items[index].name.slice(PREVIEW_PREFIX.length), settings.color))
      })
      lastPreviewSelectionKey = ''
      disableBoundaryFill(model, settings)
    } catch (error) {
      created.forEach(discardShape)
      throw error
    } finally { unselectCurrentShapes(model.app) }
    return generationSummary(model, items.map(item => item.name))
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

    const runId = Date.now().toString(36)
    const names = []
    const created = []
    try {
      for (let i = 0; i < model.regions.length; i += 1) {
        const color = settings.palette
          ? paletteColor(settings.color, i, model.regions.length)
          : settings.color
        const name = PREFIX + runId + '_' + String(i + 1).padStart(2, '0')
        created.push(createFreeform(model.slide, model.regions[i], color, settings.opacity, name))
        names.push(name)
      }
      disableBoundaryFill(model, settings)
    } catch (error) {
      created.forEach(discardShape)
      throw error
    } finally {
      unselectCurrentShapes(model.app)
    }
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

    const runId = Date.now().toString(36)
    const names = []
    const created = []
    try {
      for (let i = 0; i < model.regions.length; i += 1) {
        const key = runId + '_' + String(i + 1).padStart(2, '0')
        const name = PREVIEW_PREFIX + key
        created.push(createFreeform(model.slide, model.regions[i], settings.color, 18, name))
        created.push(createSelectionMarker(model.slide, model.regions[i], key, settings.color))
        names.push(name)
      }
      disableBoundaryFill(model, settings)
    } catch (error) {
      created.forEach(discardShape)
      clearPreviewShapes(model.slide)
      throw error
    } finally {
      unselectCurrentShapes(model.app)
    }
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

    // 若用户刚点完定位点便立即点击任务窗格，轮询可能还未来得及记录；
    // 在提交前再收集一次当前标记，确保这次点击不会丢失。
    const currentlySelected = getSelectedShapes(app, '', true).filter(isMarkerShape)
    currentlySelected.forEach(marker => {
      const preview = findPreviewByKey(slide, markerSelectionKey(marker))
      if (preview && !isPickedPreviewShape(preview)) {
        setPreviewPicked(preview, true, settings.color)
        setMarkerPicked(marker, true, settings.color)
      }
    })
    const previews = []
    for (let i = 1; i <= toNumber(slide.Shapes.Count, 0); i += 1) {
      const shape = slide.Shapes.Item(i)
      if (isPickedPreviewShape(shape)) previews.push(shape)
    }
    if (!previews.length) {
      throw new Error('还没有点选候选区域。请点击区域内部的圆形定位点；已选区域会自动加深。')
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
      if (isPreviewShape(shape) || isMarkerShape(shape)) shape.Delete()
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
    MARKER_PREFIX,
    PICKED_MARKER_PREFIX,
    DEBUG_PREFIX,
    refreshGeometrySnapshot,
    setGeometrySnapshot,
    analyze,
    generate,
    prepareManualSelection,
    generateNative: options => generateNative(options, false),
    prepareManualSelectionNative: options => generateNative(options, true),
    captureManualSelection,
    generateSelected,
    generateGeometryDebug,
    inspectSelectedGeometry,
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
      adjustedNormalizedPolygon,
      sampleFreeformPath,
      sampleCurvedConnectorPath,
      genericPathsFromShape,
      cubicSubcurve,
      freeformCommands,
      addCurvePointRows,
      nativePaths,
      ellipsePointOnRay,
      polygonArea,
      polygonCentroid,
      polygonInteriorPoint,
      segmentEllipseRoots,
      pointInPolygon
    }
  }

  global.WpsAutoFill = api
  if (typeof module !== 'undefined' && module.exports) module.exports = api
})(typeof window !== 'undefined' ? window : globalThis)
