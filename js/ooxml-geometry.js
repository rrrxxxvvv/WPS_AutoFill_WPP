(function (global) {
  'use strict'

  const ANGLE_UNIT = 60000
  const FULL_CIRCLE = 360 * ANGLE_UNIT
  const PRESET_DEFINITIONS_URL = '../assets/presetShapeDefinitions.xml'
  let presetDocumentPromise = null

  function nodeLocalName(node) {
    if (!node) return ''
    return String(node.localName || node.nodeName || '').replace(/^.*:/, '')
  }

  function elementChildren(node, name) {
    const result = []
    if (!node || !node.childNodes) return result
    for (let i = 0; i < node.childNodes.length; i += 1) {
      const child = node.childNodes[i]
      if (child.nodeType === 1 && (!name || nodeLocalName(child) === name)) result.push(child)
    }
    return result
  }

  function firstChild(node, name) {
    const children = elementChildren(node, name)
    return children.length ? children[0] : null
  }

  function descendants(node, name) {
    const result = []
    if (!node || !node.getElementsByTagName) return result
    const all = node.getElementsByTagName('*')
    for (let i = 0; i < all.length; i += 1) {
      if (nodeLocalName(all[i]) === name) result.push(all[i])
    }
    return result
  }

  function firstDescendant(node, name) {
    const items = descendants(node, name)
    return items.length ? items[0] : null
  }

  function attribute(node, name, fallback) {
    if (!node || !node.getAttribute) return fallback
    const value = node.getAttribute(name)
    return value === null || value === undefined || value === '' ? fallback : value
  }

  function xmlParser() {
    if (typeof global.DOMParser === 'function') return new global.DOMParser()
    if (typeof require === 'function') {
      try {
        const DOMParser = require('@xmldom/xmldom').DOMParser
        return new DOMParser()
      } catch (_) {}
    }
    throw new Error('当前环境没有可用的 XML 解析器。')
  }

  function parseXml(text) {
    const source = String(text || '').replace(/^\uFEFF/, '')
    const document = xmlParser().parseFromString(source, 'application/xml')
    const errors = descendants(document, 'parsererror')
    if (errors.length) throw new Error('PPTX XML 解析失败。')
    return document
  }

  function toNumber(value, fallback) {
    const number = Number(value)
    return Number.isFinite(number) ? number : fallback
  }

  function resolveValue(token, env) {
    if (token === undefined || token === null || token === '') return 0
    const number = Number(token)
    if (Number.isFinite(number)) return number
    return Number.isFinite(env[token]) ? env[token] : 0
  }

  function radians(angle) {
    return angle / ANGLE_UNIT * Math.PI / 180
  }

  function angleUnits(value) {
    return value * 180 / Math.PI * ANGLE_UNIT
  }

  function evaluateFormula(formula, env) {
    const parts = String(formula || '').trim().split(/\s+/)
    const op = parts.shift() || 'val'
    const a = resolveValue(parts[0], env)
    const b = resolveValue(parts[1], env)
    const c = resolveValue(parts[2], env)
    switch (op) {
      case 'val': return a
      case '*/': return Math.abs(c) < 1e-15 ? 0 : a * b / c
      case '+-': return a + b - c
      case '+/': return Math.abs(c) < 1e-15 ? 0 : (a + b) / c
      case '?:': return a > 0 ? b : c
      case 'abs': return Math.abs(a)
      case 'at2': return angleUnits(Math.atan2(b, a))
      case 'cat2': return a * Math.cos(Math.atan2(c, b))
      case 'cos': return a * Math.cos(radians(b))
      case 'max': return Math.max(a, b)
      case 'min': return Math.min(a, b)
      case 'mod': return Math.sqrt(a * a + b * b + c * c)
      case 'pin': return Math.max(a, Math.min(b, c))
      case 'sat2': return a * Math.sin(Math.atan2(c, b))
      case 'sin': return a * Math.sin(radians(b))
      case 'sqrt': return Math.sqrt(Math.max(0, a))
      case 'tan': return a * Math.tan(radians(b))
      default: return resolveValue(formula, env)
    }
  }

  function baseEnvironment(width, height) {
    const w = Math.max(1e-9, Math.abs(toNumber(width, 1)))
    const h = Math.max(1e-9, Math.abs(toNumber(height, 1)))
    const ss = Math.min(w, h)
    const env = {
      w, h, l: 0, t: 0, r: w, b: h,
      hc: w / 2, vc: h / 2,
      ss, ls: Math.max(w, h),
      cd2: FULL_CIRCLE / 2,
      cd4: FULL_CIRCLE / 4,
      cd8: FULL_CIRCLE / 8,
      '3cd4': FULL_CIRCLE * 3 / 4
    }
    ;[2, 3, 4, 5, 6, 8, 10, 12, 16, 32].forEach(divisor => {
      env['wd' + divisor] = w / divisor
      env['hd' + divisor] = h / divisor
      env['ssd' + divisor] = ss / divisor
    })
    return env
  }

  function evaluateGuideList(listNode, env, overrides, useOverrides) {
    if (!listNode) return
    elementChildren(listNode, 'gd').forEach(guide => {
      const name = attribute(guide, 'name', '')
      if (!name) return
      if (useOverrides && overrides && Number.isFinite(overrides[name])) {
        env[name] = overrides[name]
      } else {
        env[name] = evaluateFormula(attribute(guide, 'fmla', 'val 0'), env)
      }
    })
  }

  function adjustmentValues(geometryNode, overrides) {
    const result = Object.assign({}, overrides || {})
    const avList = firstChild(geometryNode, 'avLst')
    elementChildren(avList, 'gd').forEach(guide => {
      const name = attribute(guide, 'name', '')
      if (!name || Number.isFinite(result[name])) return
      const formula = attribute(guide, 'fmla', 'val 0').trim().split(/\s+/)
      const value = Number(formula[formula.length - 1])
      if (Number.isFinite(value)) result[name] = value
    })
    return result
  }

  function point(x, y) {
    return { x: toNumber(x, 0), y: toNumber(y, 0) }
  }

  function cubicPoint(cubic, t) {
    const mt = 1 - t
    const mt2 = mt * mt
    const t2 = t * t
    return point(
      cubic.p0.x * mt2 * mt + 3 * cubic.p1.x * mt2 * t + 3 * cubic.p2.x * mt * t2 + cubic.p3.x * t2 * t,
      cubic.p0.y * mt2 * mt + 3 * cubic.p1.y * mt2 * t + 3 * cubic.p2.y * mt * t2 + cubic.p3.y * t2 * t
    )
  }

  function transformForShape(shape) {
    const left = toNumber(shape && shape.Left, 0)
    const top = toNumber(shape && shape.Top, 0)
    const width = Math.abs(toNumber(shape && shape.Width, 0))
    const height = Math.abs(toNumber(shape && shape.Height, 0))
    const cx = left + width / 2
    const cy = top + height / 2
    const angle = toNumber(shape && shape.Rotation, 0) * Math.PI / 180
    const cos = Math.cos(angle)
    const sin = Math.sin(angle)
    let hFlip = false
    let vFlip = false
    try { hFlip = Number(shape.HorizontalFlip) === -1 || shape.HorizontalFlip === true } catch (_) {}
    try { vFlip = Number(shape.VerticalFlip) === -1 || shape.VerticalFlip === true } catch (_) {}
    return {
      width,
      height,
      point(local) {
        let x = left + (hFlip ? width - local.x : local.x)
        let y = top + (vFlip ? height - local.y : local.y)
        if (Math.abs(angle) > 1e-12) {
          const dx = x - cx
          const dy = y - cy
          x = cx + dx * cos - dy * sin
          y = cy + dx * sin + dy * cos
        }
        return point(x, y)
      }
    }
  }

  function arcCubics(current, rx, ry, startAngle, sweepAngle) {
    // DrawingML angles describe a ray from the ellipse centre, not the
    // parameter in (rx*cos(t), ry*sin(t)). Preserve winding across +/-pi.
    const parameterAngle = angle => {
      const turns = Math.floor((angle + Math.PI) / (2 * Math.PI))
      const principal = angle - turns * 2 * Math.PI
      return Math.atan2(rx * Math.sin(principal), ry * Math.cos(principal)) + turns * 2 * Math.PI
    }
    const polarStart = radians(startAngle)
    const start = parameterAngle(polarStart)
    const sweep = parameterAngle(polarStart + radians(sweepAngle)) - start
    // 使用不超过 45° 的三次贝塞尔子段逼近 OOXML 椭圆弧；相较常见的
    // 90° 单段方案，交点和最终边界的径向误差降到视觉与拓扑容差以下。
    const chunks = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 4)))
    const delta = sweep / chunks
    const center = point(current.x - rx * Math.cos(start), current.y - ry * Math.sin(start))
    const result = []
    for (let index = 0; index < chunks; index += 1) {
      const a0 = start + delta * index
      const a1 = a0 + delta
      const k = 4 / 3 * Math.tan(delta / 4)
      const p0 = point(center.x + rx * Math.cos(a0), center.y + ry * Math.sin(a0))
      const p3 = point(center.x + rx * Math.cos(a1), center.y + ry * Math.sin(a1))
      result.push({
        p0,
        p1: point(p0.x - rx * Math.sin(a0) * k, p0.y + ry * Math.cos(a0) * k),
        p2: point(p3.x + rx * Math.sin(a1) * k, p3.y - ry * Math.cos(a1) * k),
        p3
      })
    }
    return result
  }

  function pushDistinct(points, value, epsilon) {
    if (!points.length || Math.hypot(points[points.length - 1].x - value.x, points[points.length - 1].y - value.y) > epsilon) {
      points.push(value)
      return true
    }
    return false
  }

  function buildSampledPath(commands, closed, transform, segments, style) {
    if (!commands.length) return null
    const points = []
    const edgeMeta = []
    const epsilon = Math.max(1e-8, Math.max(transform.width, transform.height, 1) * 1e-9)
    const first = transform.point(commands[0].p0)
    pushDistinct(points, first, epsilon)
    commands.forEach(command => {
      if (command.kind === 'line') {
        const destination = transform.point(command.p1)
        if (pushDistinct(points, destination, epsilon)) edgeMeta.push({ kind: 'line' })
        return
      }
      const cubic = {
        p0: transform.point(command.cubic.p0),
        p1: transform.point(command.cubic.p1),
        p2: transform.point(command.cubic.p2),
        p3: transform.point(command.cubic.p3)
      }
      const steps = Math.max(4, Math.min(96, Math.ceil(toNumber(segments, 160) / Math.max(1, commands.length))))
      for (let step = 1; step <= steps; step += 1) {
        const before = points.length
        pushDistinct(points, cubicPoint(cubic, step / steps), epsilon)
        if (points.length > before) {
          edgeMeta.push({ kind: 'curve', curve: cubic, t0: (step - 1) / steps, t1: step / steps })
        }
      }
    })
    if (closed && points.length > 1) {
      const last = points[points.length - 1]
      if (Math.hypot(last.x - first.x, last.y - first.y) <= epsilon) points.pop()
    }
    if (points.length < (closed ? 3 : 2)) return null
    if (edgeMeta.length !== (closed ? points.length : points.length - 1)) {
      while (edgeMeta.length < (closed ? points.length : points.length - 1)) edgeMeta.push({ kind: 'line' })
      if (edgeMeta.length > (closed ? points.length : points.length - 1)) edgeMeta.length = closed ? points.length : points.length - 1
    }
    return Object.assign({ points, edgeMeta, closed, curved: edgeMeta.some(edge => edge.kind === 'curve') }, style || {})
  }

  function evaluatePathList(geometryNode, env, shape, segments) {
    const pathList = firstChild(geometryNode, 'pathLst')
    if (!pathList) return []
    const transform = transformForShape(shape)
    const result = []
    elementChildren(pathList, 'path').forEach(pathNode => {
      const pathWidth = toNumber(attribute(pathNode, 'w', ''), 0)
      const pathHeight = toNumber(attribute(pathNode, 'h', ''), 0)
      const scaleX = pathWidth > 0 ? transform.width / pathWidth : 1
      const scaleY = pathHeight > 0 ? transform.height / pathHeight : 1
      const valuePoint = ptNode => point(
        resolveValue(attribute(ptNode, 'x', '0'), env) * scaleX,
        resolveValue(attribute(ptNode, 'y', '0'), env) * scaleY
      )
      const valueX = token => resolveValue(token, env) * scaleX
      const valueY = token => resolveValue(token, env) * scaleY
      let start = null
      let current = null
      let commands = []
      let sawClose = false
      const finish = () => {
        if (!commands.length || !start) {
          start = null
          current = null
          commands = []
          sawClose = false
          return
        }
        const sampled = buildSampledPath(commands, sawClose && commands.length >= 2, transform, segments, {
          fill: attribute(pathNode, 'fill', 'norm'),
          stroke: attribute(pathNode, 'stroke', 'true') !== 'false'
        })
        if (sampled) result.push(sampled)
        start = null
        current = null
        commands = []
        sawClose = false
      }
      elementChildren(pathNode).forEach(commandNode => {
        const kind = nodeLocalName(commandNode)
        if (kind === 'moveTo') {
          finish()
          const pt = firstChild(commandNode, 'pt')
          start = current = valuePoint(pt)
          return
        }
        if (!current) return
        if (kind === 'lnTo') {
          const destination = valuePoint(firstChild(commandNode, 'pt'))
          commands.push({ kind: 'line', p0: current, p1: destination })
          current = destination
        } else if (kind === 'cubicBezTo') {
          const points = elementChildren(commandNode, 'pt').map(valuePoint)
          if (points.length === 3) {
            const cubic = { p0: current, p1: points[0], p2: points[1], p3: points[2] }
            commands.push({ kind: 'cubic', cubic, p0: current, p1: points[2] })
            current = points[2]
          }
        } else if (kind === 'quadBezTo') {
          const points = elementChildren(commandNode, 'pt').map(valuePoint)
          if (points.length === 2) {
            const control = points[0]
            const end = points[1]
            const cubic = {
              p0: current,
              p1: point(current.x + (control.x - current.x) * 2 / 3, current.y + (control.y - current.y) * 2 / 3),
              p2: point(end.x + (control.x - end.x) * 2 / 3, end.y + (control.y - end.y) * 2 / 3),
              p3: end
            }
            commands.push({ kind: 'cubic', cubic, p0: current, p1: end })
            current = end
          }
        } else if (kind === 'arcTo') {
          const cubics = arcCubics(
            current,
            Math.abs(valueX(attribute(commandNode, 'wR', '0'))),
            Math.abs(valueY(attribute(commandNode, 'hR', '0'))),
            resolveValue(attribute(commandNode, 'stAng', '0'), env),
            resolveValue(attribute(commandNode, 'swAng', '0'), env)
          )
          cubics.forEach(cubic => {
            commands.push({ kind: 'cubic', cubic, p0: cubic.p0, p1: cubic.p3 })
            current = cubic.p3
          })
        } else if (kind === 'close' && start) {
          if (Math.hypot(current.x - start.x, current.y - start.y) > 1e-9) {
            commands.push({ kind: 'line', p0: current, p1: start })
            current = start
          }
          sawClose = true
        }
      })
      finish()
    })
    return result
  }

  function evaluateGeometry(geometryNode, adjustments, shape, segments) {
    if (!geometryNode) return []
    const transform = transformForShape(shape)
    const env = baseEnvironment(transform.width, transform.height)
    const values = adjustmentValues(geometryNode, adjustments)
    Object.keys(values).forEach(name => { env[name] = values[name] })
    evaluateGuideList(firstChild(geometryNode, 'avLst'), env, values, true)
    evaluateGuideList(firstChild(geometryNode, 'gdLst'), env, null, false)
    return evaluatePathList(geometryNode, env, shape, segments)
  }

  function parseAdjustments(prstGeometry) {
    const result = {}
    const list = firstChild(prstGeometry, 'avLst')
    elementChildren(list, 'gd').forEach(guide => {
      const name = attribute(guide, 'name', '')
      const formula = attribute(guide, 'fmla', '').trim().split(/\s+/)
      const value = Number(formula[formula.length - 1])
      if (name && Number.isFinite(value)) result[name] = value
    })
    return result
  }

  function parseTransform(shapeNode) {
    const shapeProperties = firstChild(shapeNode, 'spPr') || firstDescendant(shapeNode, 'spPr')
    const transform = firstChild(shapeProperties, 'xfrm')
    const offset = firstChild(transform, 'off')
    const extent = firstChild(transform, 'ext')
    if (!transform || !offset || !extent) return null
    return {
      x: toNumber(attribute(offset, 'x', '0'), 0),
      y: toNumber(attribute(offset, 'y', '0'), 0),
      cx: toNumber(attribute(extent, 'cx', '0'), 0),
      cy: toNumber(attribute(extent, 'cy', '0'), 0),
      rotation: toNumber(attribute(transform, 'rot', '0'), 0),
      horizontalFlip: attribute(transform, 'flipH', '0') === '1',
      verticalFlip: attribute(transform, 'flipV', '0') === '1'
    }
  }

  function geometryEntries(slideXml) {
    const document = parseXml(slideXml)
    const result = {}
    ;['sp', 'cxnSp'].forEach(tag => {
      descendants(document, tag).forEach(shapeNode => {
        let ancestor = shapeNode.parentNode
        let inGroup = false
        while (ancestor && ancestor !== document) {
          if (nodeLocalName(ancestor) === 'grpSp') {
            inGroup = true
            break
          }
          ancestor = ancestor.parentNode
        }
        const properties = firstDescendant(shapeNode, 'cNvPr')
        const id = attribute(properties, 'id', '')
        if (!id) return
        const prstGeometry = firstDescendant(shapeNode, 'prstGeom')
        const customGeometry = firstDescendant(shapeNode, 'custGeom')
        if (!prstGeometry && !customGeometry) return
        result[id] = {
          id,
          name: attribute(properties, 'name', ''),
          preset: prstGeometry ? attribute(prstGeometry, 'prst', '') : '',
          adjustments: prstGeometry ? parseAdjustments(prstGeometry) : parseAdjustments(customGeometry),
          transform: parseTransform(shapeNode),
          inGroup,
          customGeometryXml: customGeometry && typeof global.XMLSerializer === 'function'
            ? new global.XMLSerializer().serializeToString(customGeometry)
            : customGeometry && customGeometry.toString ? customGeometry.toString() : ''
        }
      })
    })
    return result
  }

  function relationshipMap(xml) {
    const document = parseXml(xml)
    const result = {}
    descendants(document, 'Relationship').forEach(rel => {
      result[attribute(rel, 'Id', '')] = attribute(rel, 'Target', '')
    })
    return result
  }

  function normalizeZipPath(base, target) {
    const parts = (base + '/' + target).replace(/\\/g, '/').split('/')
    const output = []
    parts.forEach(part => {
      if (!part || part === '.') return
      if (part === '..') output.pop()
      else output.push(part)
    })
    return output.join('/')
  }

  async function slidePartPath(zip, slideIndex) {
    const presentationXml = await zip.file('ppt/presentation.xml').async('string')
    const relsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
    const document = parseXml(presentationXml)
    const slideIds = descendants(document, 'sldId')
    const target = slideIds[Math.max(0, slideIndex - 1)]
    if (!target) throw new Error('无法在 PPTX 中定位当前幻灯片。')
    const relationshipId = attribute(target, 'r:id', '') || attribute(target, 'id', '')
    const path = relationshipMap(relsXml)[relationshipId]
    if (!path) throw new Error('当前幻灯片的 PPTX 关系缺失。')
    return normalizeZipPath('ppt', path)
  }

  function presetMap(document) {
    const result = {}
    const root = document && document.documentElement
    elementChildren(root).forEach(node => { result[nodeLocalName(node)] = node })
    return result
  }

  async function loadPresetDocument() {
    if (!presetDocumentPromise) {
      presetDocumentPromise = (async () => {
        if (typeof fetch !== 'function') throw new Error('当前环境无法读取标准预设形状定义。')
        const response = await fetch(PRESET_DEFINITIONS_URL)
        if (!response.ok) throw new Error('标准预设形状定义读取失败：HTTP ' + response.status)
        return parseXml(await response.text())
      })()
    }
    return presetDocumentPromise
  }

  function setPresetDefinitions(xmlText) {
    presetDocumentPromise = Promise.resolve(parseXml(xmlText))
  }

  async function snapshotFromBuffer(buffer, slideIndex, presetXmlText) {
    if (!global.JSZip || typeof global.JSZip.loadAsync !== 'function') {
      throw new Error('PPTX 解包组件尚未加载。')
    }
    const zip = await global.JSZip.loadAsync(buffer)
    const partPath = await slidePartPath(zip, slideIndex)
    const slidePart = zip.file(partPath)
    if (!slidePart) throw new Error('PPTX 中不存在当前幻灯片部件：' + partPath)
    const definitions = presetXmlText ? parseXml(presetXmlText) : await loadPresetDocument()
    const shapes = geometryEntries(await slidePart.async('string'))
    const presets = presetMap(definitions)
    const pathCache = {}
    return {
      partPath,
      shapes,
      createNativePackage(items, width, height) {
        if (!global.WpsNativeShapeWriter) throw new Error('原生形状写入组件未加载。')
        return global.WpsNativeShapeWriter.createPackage(items, width, height, buffer)
      },
      geometryForShape(shape, segments) {
        let id = ''
        try { id = String(shape.Id) } catch (_) {}
        const entry = shapes[id]
        if (!entry) return []
        const cacheKey = entry.inGroup
          ? [id, segments, shape.Left, shape.Top, shape.Width, shape.Height, shape.Rotation, shape.HorizontalFlip, shape.VerticalFlip].join('|')
          : [id, segments].join('|')
        if (pathCache[cacheKey]) return pathCache[cacheKey].map(path => Object.assign({}, path, { shape }))
        let geometry = entry.customGeometryXml
          ? parseXml(entry.customGeometryXml).documentElement
          : presets[entry.preset]
        let geometryShape = shape
        if (entry.transform && !entry.inGroup && entry.transform.cx > 0 && entry.transform.cy > 0) {
          geometryShape = {
            Id: shape.Id,
            Left: entry.transform.x / 12700,
            Top: entry.transform.y / 12700,
            Width: entry.transform.cx / 12700,
            Height: entry.transform.cy / 12700,
            Rotation: entry.transform.rotation / ANGLE_UNIT,
            HorizontalFlip: entry.transform.horizontalFlip,
            VerticalFlip: entry.transform.verticalFlip
          }
        }
        // ECMA-376 第一版附带的预设定义以 downArrow 作为竖向箭头母版，
        // 但 PPTX 仍会写入 prst="upArrow"。缺失时按标准镜像母版处理。
        if (!geometry && entry.preset === 'upArrow' && presets.downArrow) {
          geometry = presets.downArrow
          geometryShape = {
            Id: shape.Id,
            Left: geometryShape.Left,
            Top: geometryShape.Top,
            Width: geometryShape.Width,
            Height: geometryShape.Height,
            Rotation: geometryShape.Rotation,
            HorizontalFlip: geometryShape.HorizontalFlip,
            VerticalFlip: !(Number(geometryShape.VerticalFlip) === -1 || geometryShape.VerticalFlip === true)
          }
        }
        const evaluated = evaluateGeometry(geometry, entry.adjustments, geometryShape, segments)
          .map(path => Object.assign(path, { name: entry.name, preset: entry.preset, source: 'pptx-ooxml' }))
        pathCache[cacheKey] = evaluated
        return evaluated.map(path => Object.assign({}, path, { shape }))
      }
    }
  }

  function joinPath(directory, filename) {
    const separator = /\\/.test(directory) ? '\\' : '/'
    return String(directory || '').replace(/[\\/]+$/g, '') + separator + filename
  }

  function binaryStringBuffer(value) {
    const text = String(value || '')
    const bytes = new Uint8Array(text.length)
    for (let i = 0; i < text.length; i += 1) bytes[i] = text.charCodeAt(i) & 0xff
    return bytes.buffer
  }

  async function refreshFromApplication(app) {
    const wps = global.wps
    if (!wps || !wps.FileSystem) throw new Error('当前 WPS 未提供文件读取接口，无法取得真实形状几何。')
    const presentation = app && app.ActivePresentation
    if (!presentation) throw new Error('当前没有打开的演示文稿。')
    let slideIndex = 1
    try { slideIndex = Number(app.ActiveWindow.View.Slide.SlideIndex) || 1 } catch (_) {}
    const fileSystem = wps.FileSystem
    const tempDirectory = wps.Env && typeof wps.Env.GetTempPath === 'function'
      ? wps.Env.GetTempPath()
      : (typeof fileSystem.tmpdir === 'function' ? fileSystem.tmpdir() : '')
    const tempPath = joinPath(tempDirectory, 'WPS_AutoFill_Geometry_' + Date.now().toString(36) + '.pptx')
    let sourcePath = ''
    let temporary = false
    try {
      if (typeof presentation.SaveCopyAs === 'function' && tempDirectory) {
        try {
          presentation.SaveCopyAs(tempPath, 24)
        } catch (_) {
          // 部分 WPS 构建只接受 FileName；保留一次兼容调用，随后仍由
          // JSZip 校验副本确实是 Open XML 包，绝不把旧版二进制当作几何源。
          presentation.SaveCopyAs(tempPath)
        }
        sourcePath = tempPath
        temporary = true
      } else {
        sourcePath = String(presentation.FullName || '')
      }
      if (!sourcePath) throw new Error('演示文稿尚未保存，无法读取 PPTX 几何。')
      let buffer = null
      if (typeof fileSystem.ReadFileAsArrayBuffer === 'function') {
        try { buffer = fileSystem.ReadFileAsArrayBuffer(sourcePath) } catch (_) { buffer = null }
      }
      if (!buffer && typeof fileSystem.readAsBinaryString === 'function') {
        buffer = binaryStringBuffer(fileSystem.readAsBinaryString(sourcePath))
      }
      if (buffer && typeof buffer.then === 'function') buffer = await buffer
      if (!buffer) throw new Error('WPS 文件接口返回了空数据。')
      return await snapshotFromBuffer(buffer, slideIndex)
    } catch (error) {
      const message = String(error && error.message ? error.message : error)
      throw new Error('无法读取当前演示文稿的真实 PPTX 几何：' + message)
    } finally {
      if (temporary) {
        try {
          if (typeof fileSystem.unlinkSync === 'function') fileSystem.unlinkSync(tempPath)
          else if (typeof fileSystem.Remove === 'function') fileSystem.Remove(tempPath)
        } catch (_) {}
      }
    }
  }

  const api = {
    refreshFromApplication,
    snapshotFromBuffer,
    setPresetDefinitions,
    parseXml,
    geometryEntries,
    evaluateFormula,
    evaluateGeometry,
    baseEnvironment,
    _internals: { elementChildren, descendants, arcCubics, normalizeZipPath, parseAdjustments, parseTransform, binaryStringBuffer }
  }

  global.WpsOoxmlGeometry = api
  if (typeof module !== 'undefined' && module.exports) module.exports = api
})(typeof window !== 'undefined' ? window : globalThis)
