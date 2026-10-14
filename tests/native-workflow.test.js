const assert = require('assert')
const fs = require('fs')
global.JSZip = require('jszip')
global.DOMParser = require('@xmldom/xmldom').DOMParser
const writer = require('../js/native-shape-writer')
const geometry = require('../js/ooxml-geometry')
const engine = require('../js/fill-engine')

async function scenario(actions, failure) {
  const inputs = [
    { name: 'curved source', paths: [{ start: { x: 100, y: 100 }, commands: [
      { kind: 'curve', p1: { x: 100, y: 200 }, p2: { x: 200, y: 200 }, end: { x: 200, y: 100 } },
      { kind: 'line', end: { x: 100, y: 100 } }
    ] }] },
    { name: 'straight source', paths: [{ start: { x: 300, y: 100 }, commands: [
      { kind: 'line', end: { x: 400, y: 100 } }, { kind: 'line', end: { x: 350, y: 200 } },
      { kind: 'line', end: { x: 300, y: 100 } }
    ] }] }
  ].map(item => ({ ...item, color: '#407cff', opacity: 100 }))
  const initial = await writer.createPackage(inputs, 960, 540, fs.readFileSync('tests/fixtures/adjusted-shapes.pptx'))
  const snapshot = await geometry.snapshotFromBuffer(initial, 1, '<presets/>')
  const shapes = []
  let nextId = 100, sourceEntries, pasteCount = 0, closed = 0, removed = 0
  const placement = entry => ({ Name: entry.name, Left: entry.transform.x / 12700,
    Top: entry.transform.y / 12700, Width: entry.transform.cx / 12700,
    Height: entry.transform.cy / 12700, Rotation: 0 })
  function add(properties) {
    const shape = { Id: nextId++, Type: 5, Connector: 0, Rotation: 0,
      Fill: { Visible: -1, ForeColor: {}, Solid() {} }, Line: { ForeColor: {} },
      ZOrder() {}, Delete() { const i = shapes.indexOf(shape); if (i >= 0) shapes.splice(i, 1) }, ...properties }
    shapes.push(shape)
    return shape
  }
  Object.values(snapshot.shapes).forEach(entry => add({ ...placement(entry), Id: Number(entry.id) }))
  const originals = shapes.slice()
  const slide = { Shapes: {
    get Count() { return shapes.length }, Item: i => shapes[i - 1],
    AddShape(type, left, top, width, height) { return add({ Type: 1, AutoShapeType: type, Left: left, Top: top, Width: width, Height: height }) },
    Paste() {
      pasteCount++
      // Simulate WPS cascading paste when previous preview/output shapes exist.
      const offset = shapes.length > originals.length ? 9 * pasteCount : 0
      const pasted = sourceEntries.slice().reverse().map(entry => {
        const shape = add({ ...entry, Left: entry.Left + offset, Top: entry.Top + offset })
        if (failure === 'resize' && pasteCount === 2) shape.Width += 1
        if (failure === 'locked-position' && pasteCount === 2) {
          const left = shape.Left
          Object.defineProperty(shape, 'Left', { get: () => left, set() {} })
        }
        return shape
      })
      return { Count: pasted.length, Item: i => pasted[i - 1] }
    }
  } }
  const app = { ActivePresentation: { PageSetup: { SlideWidth: 960, SlideHeight: 540 } },
    ActiveWindow: { View: { Slide: slide }, Selection: { ShapeRange: null, Unselect() {} } },
    Presentations: { Open() { return { Slides: { Item: () => ({ Shapes: {
      Count: sourceEntries.length, Item: i => sourceEntries[i - 1], Range: () => ({ Copy() {} })
    } }) }, Close() { closed++ } } } }
  }
  global.Application = app
  global.wps = { Env: { GetTempPath: () => '/mock-temp' }, FileSystem: {
    writeAsBinaryString() {}, unlinkSync() { removed++ }
  } }
  const createPackage = snapshot.createNativePackage
  snapshot.createNativePackage = async (...args) => {
    const bytes = await createPackage(...args)
    const output = await geometry.snapshotFromBuffer(bytes, 1, '<presets/>')
    sourceEntries = Object.values(output.shapes).map(placement)
    return bytes
  }
  engine.setGeometrySnapshot(snapshot)
  let previous
  for (const action of actions) {
    previous = shapes.slice()
    const run = action === 'preview' ? engine.prepareManualSelectionNative : engine.generateNative
    if (failure && pasteCount === 1) {
      await assert.rejects(run({ disableBoundaryFill: false }), /位置|尺寸/)
      assert.deepStrictEqual(shapes, previous, 'failed import must keep previous candidates and originals')
      break
    }
    const result = await run({ disableBoundaryFill: false })
    assert.strictEqual(result.count, 2)
    if (action === 'preview') {
      assert.strictEqual(shapes.filter(shape => /WPS_AutoFill_Preview_/.test(shape.Name)).length, 2)
      assert.strictEqual(shapes.filter(shape => /WPS_AutoFill_Marker_/.test(shape.Name)).length, 2)
    }
    sourceEntries.forEach(expected => {
      const actual = shapes.find(shape => shape.Name === expected.Name)
      for (const key of ['Left', 'Top', 'Width', 'Height', 'Rotation']) assert.strictEqual(actual[key], expected[key], key)
    })
  }
  assert.strictEqual(closed, pasteCount)
  assert.strictEqual(removed, pasteCount)
  originals.forEach(original => assert.ok(shapes.includes(original)))
  if (!failure && actions.at(-1) === 'fill') {
    assert.ok(!shapes.some(shape => /WPS_AutoFill_(Preview|Marker)_/.test(shape.Name)), 'all candidates and markers must be removed after fill')
  }
  return sourceEntries.map(({ Name, ...position }) => position)
}

async function main() {
  const direct = await scenario(['fill'])
  assert.deepStrictEqual(await scenario(['preview', 'fill']), direct)
  assert.deepStrictEqual(await scenario(['preview', 'preview', 'fill']), direct)
  assert.deepStrictEqual(await scenario(['fill', 'fill']), direct)
  await scenario(['preview', 'fill'], 'locked-position')
  await scenario(['preview', 'fill'], 'resize')
  console.log('Native preview → fill workflow: translation restoration, reordered paste, and rollback passed')
}
main().catch(error => { console.error(error); process.exitCode = 1 })
