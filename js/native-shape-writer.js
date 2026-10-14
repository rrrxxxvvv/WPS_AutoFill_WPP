/* Native editable DrawingML output. Avoid WPS FreeformBuilder's spline coercion. */
;(function (global) {
  'use strict'
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const xml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
  const emu = value => {
    if (!Number.isFinite(value)) throw new Error('原生形状包含无效坐标。')
    return Math.round(value * 12700)
  }

  function shapeXml(item, index) {
    const points = item.paths.flatMap(path => [path.start].concat(path.commands.flatMap(c => c.kind === 'curve' ? [c.p1, c.p2, c.end] : [c.end])))
    const left = Math.min(...points.map(p => p.x))
    const top = Math.min(...points.map(p => p.y))
    const width = Math.max(1, emu(Math.max(...points.map(p => p.x)) - left))
    const height = Math.max(1, emu(Math.max(...points.map(p => p.y)) - top))
    const pt = p => '<a:pt x="' + emu(p.x - left) + '" y="' + emu(p.y - top) + '"/>'
    const paths = '<a:path w="' + width + '" h="' + height + '">' + item.paths.map(path => '<a:moveTo>' + pt(path.start) + '</a:moveTo>' + path.commands.map(c => c.kind === 'curve'
      ? '<a:cubicBezTo>' + pt(c.p1) + pt(c.p2) + pt(c.end) + '</a:cubicBezTo>'
      : '<a:lnTo>' + pt(c.end) + '</a:lnTo>').join('') + '<a:close/>').join('') + '</a:path>'
    const color = String(item.color || '#4f7cff').replace('#', '')
    if (!/^[0-9a-f]{6}$/i.test(color)) throw new Error('原生形状颜色无效。')
    return '<p:sp><p:nvSpPr><p:cNvPr id="' + (index + 2) + '" name="' + xml(item.name) + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
      '<a:xfrm><a:off x="' + emu(left) + '" y="' + emu(top) + '"/><a:ext cx="' + width + '" cy="' + height + '"/></a:xfrm>' +
      '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>' + paths + '</a:pathLst></a:custGeom>' +
      (item.fill === false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + color + '"><a:alpha val="' + Math.round(Math.max(0, Math.min(100, item.opacity)) * 1000) + '"/></a:srgbClr></a:solidFill>') +
      (item.stroke ? '<a:ln w="12700"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></a:ln>' : '<a:ln><a:noFill/></a:ln>') + '</p:spPr></p:sp>'
  }

  async function createPackage(items, width, height, templateBuffer) {
    const zip = templateBuffer ? await global.JSZip.loadAsync(templateBuffer) : new global.JSZip()
    if (templateBuffer) {
      // Keep the source document's masters/layouts: WPS rejects an otherwise
      // valid geometry-only package without its presentation infrastructure.
      const presentation = await zip.file('ppt/presentation.xml').async('string')
      zip.file('ppt/presentation.xml', presentation.replace(/<p:sldIdLst[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst><p:sldId id="256" r:id="autofillNativeSlide"/></p:sldIdLst>'))
      const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
      zip.file('ppt/_rels/presentation.xml.rels', rels.replace('</Relationships>', '<Relationship Id="autofillNativeSlide" Type="' + R + '/slide" Target="slides/autofill-native.xml"/></Relationships>'))
      const contentTypes = await zip.file('[Content_Types].xml').async('string')
      zip.file('[Content_Types].xml', contentTypes.replace('</Types>', '<Override PartName="/ppt/slides/autofill-native.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'))
      const slideRels = Object.keys(zip.files).find(name => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(name))
      if (slideRels) zip.file('ppt/slides/_rels/autofill-native.xml.rels', await zip.file(slideRels).async('string'))
    } else {
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + R + '/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    zip.file('ppt/presentation.xml', '<p:presentation xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:p="' + P + '"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="' + emu(width) + '" cy="' + emu(height) + '"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + R + '/slide" Target="slides/slide1.xml"/></Relationships>')
    }
    zip.file(templateBuffer ? 'ppt/slides/autofill-native.xml' : 'ppt/slides/slide1.xml', '<p:sld xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:p="' + P + '"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + items.map(shapeXml).join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  }

  async function importPackage(app, slide, bytes, items) {
    const fs = global.wps && global.wps.FileSystem
    if (!fs || typeof fs.writeAsBinaryString !== 'function') throw new Error('当前 WPS 缺少原生形状临时文件写入接口。')
    const directory = global.wps.Env && global.wps.Env.GetTempPath ? global.wps.Env.GetTempPath() : fs.tmpdir()
    if (!directory) throw new Error('无法取得 WPS 临时目录。')
    const filename = String(directory).replace(/[\\/]+$/, '') + '/WPS_AutoFill_Native_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2) + '.pptx'
    const chunks = []
    for (let i = 0; i < bytes.length; i += 8192) chunks.push(String.fromCharCode.apply(null, bytes.subarray(i, i + 8192)))
    let source = null
    const created = []
    const existingIds = new Set()
    for (let i = 1; i <= Number(slide.Shapes.Count); i++) existingIds.add(String(slide.Shapes.Item(i).Id))
    try {
      await fs.writeAsBinaryString(filename, chunks.join(''))
      source = app.Presentations.Open(filename, -1, 0, 0)
      if (!source) throw new Error('WPS 未能打开原生曲边临时文稿。')
      const sourceShapes = source.Slides.Item(1).Shapes
      const sourcePositions = new Map()
      const readPlacement = shape => {
        const placement = {}
        for (const key of ['Left', 'Top', 'Width', 'Height', 'Rotation']) {
          const value = shape[key]
          if (value === null || value === undefined || !Number.isFinite(Number(value))) {
            throw new Error('WPS 未能返回原生色块的位置：' + key)
          }
          placement[key] = Number(value)
        }
        return placement
      }
      for (let i = 1; i <= Number(sourceShapes.Count); i++) {
        const shape = sourceShapes.Item(i)
        const name = String(shape.Name)
        if (sourcePositions.has(name)) throw new Error('原生色块名称重复，无法校准位置。')
        sourcePositions.set(name, readPlacement(shape))
      }
      if (sourcePositions.size !== items.length || items.some(item => !sourcePositions.has(item.name))) {
        throw new Error('原生色块与待生成区域不一致。')
      }
      sourceShapes.Range().Copy()
      const pasted = slide.Shapes.Paste()
      if (pasted) for (let i = 1; i <= Number(pasted.Count); i++) created.push(pasted.Item(i))
      if (created.length !== items.length) throw new Error('WPS 导入形状数量与区域数不一致。')
      const pastedNames = new Set()
      created.forEach(shape => {
        const name = String(shape.Name)
        const expected = sourcePositions.get(name)
        if (!expected || pastedNames.has(name)) throw new Error('WPS 粘贴后色块身份不一致，无法校准位置。')
        pastedNames.add(name)
        const actual = readPlacement(shape)
        const tolerance = 0.02 // points; float precision, not a visual snapping tolerance
        if (['Width', 'Height', 'Rotation'].some(key => Math.abs(actual[key] - expected[key]) > tolerance)) {
          throw new Error('WPS 粘贴改变了色块尺寸或旋转，已停止生成。')
        }
        // Repeated paste may apply a cascade translation while preview shapes
        // still exist. Restore each source position by NAME, never by range order
        // or by fitting the curve's control-point bounding box. Do not resize.
        shape.Left = expected.Left
        shape.Top = expected.Top
        const restored = readPlacement(shape)
        if (Object.keys(expected).some(key => Math.abs(restored[key] - expected[key]) > tolerance)) {
          throw new Error('WPS 未能恢复色块的原始位置，已停止生成。')
        }
      })
      created.forEach(shape => shape.ZOrder(1))
      return created
    } catch (error) {
      // Paste may rename a shape. The returned objects still belong to this
      // transaction, so rollback must not rely on their names alone.
      created.forEach(shape => {
        if (!existingIds.has(String(shape.Id))) { try { shape.Delete() } catch (_) {} }
      })
      for (let i = Number(slide.Shapes.Count); i >= 1; i--) {
        const shape = slide.Shapes.Item(i)
        if (!existingIds.has(String(shape.Id)) && items.some(item => item.name === String(shape.Name))) {
          try { shape.Delete() } catch (_) {}
        }
      }
      throw error
    } finally {
      if (source) { try { source.Close() } catch (_) {} }
      try { if (fs.unlinkSync) fs.unlinkSync(filename); else if (fs.Remove) fs.Remove(filename) } catch (_) {}
    }
  }

  const api = { createPackage, shapeXml, importPackage }
  global.WpsNativeShapeWriter = api
  if (typeof module !== 'undefined' && module.exports) module.exports = api
})(typeof window !== 'undefined' ? window : globalThis)
