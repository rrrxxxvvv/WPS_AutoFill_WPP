const assert = require('assert')
const fs = require('fs')
global.JSZip = require('jszip')
const geometry = require('../js/ooxml-geometry')
require('../js/native-shape-writer')
const g = require('../js/fill-engine')._geometry

// Reconstructed from the can + bevel screenshot. The can's left ellipse seam
// is 1.125 units below the bevel's top edge: they must NOT be snapped together.
async function main() {
  const template = fs.readFileSync('tests/fixtures/adjusted-shapes.pptx')
  const definitions = fs.readFileSync('assets/presetShapeDefinitions.xml', 'utf8')
  let probe = null
  for (const scale of [0.25, 0.5, 1, 2]) {
    for (const seamOffset of [-1.125, 0, 0.05, 1.125, 2]) {
      const snapshot = await geometry.snapshotFromBuffer(template, 2, definitions)
      const specs = [
        ['b', 'bevel', 74, 131, 636, 354],
        ['c', 'can', 490, 131 - 433 / 8 + seamOffset, 433, 514]
      ]
      specs.forEach(([id, preset, x, y, cx, cy]) => {
        snapshot.shapes[id] = { id, name: preset, preset, adjustments: {}, transform: {
          x: x * scale * 12700, y: y * scale * 12700,
          cx: cx * scale * 12700, cy: cy * scale * 12700, rotation: 0
        } }
      })
      const paths = specs.flatMap(([Id]) => snapshot.geometryForShape({ Id }, 160).filter(p => p.stroke !== false))
      const regions = g.buildGenericRegions(paths)
      const picked = regions.filter(p => g.pointInPolygon({ x: 293 * scale, y: 153 * scale }, p))
      assert.strictEqual(picked.length, 1, 'left upper point belongs to exactly one face')
      assert.ok(Math.max(...picked[0].map(p => p.x)) < 491 * scale,
        'left upper face must stop at cylinder wall, not continue to bevel right corner')
      assert.ok(!g.pointInPolygon({ x: 640 * scale, y: 153 * scale }, picked[0]),
        'upper right and upper left faces must not be joined')
      for (const region of regions) {
        const vertices = region.map(p => p.x.toFixed(7) + ':' + p.y.toFixed(7))
        assert.strictEqual(new Set(vertices).size, vertices.length,
          'each face is a simple cycle, not a retracing walk through merged junctions')
        const output = g.nativePaths(region)[0]
        assert.ok(output.commands.length < 30, 'retain trimmed cubics, not dense polyline output')
      }
      if (scale === 1 && seamOffset === 1.125) {
        assert.strictEqual(regions.length, 13, 'screenshot reconstruction has 13 faces, previously 11')
        const items = regions.map((p, i) => ({ name: 'REGION_' + i, paths: g.nativePaths(p), color: '#4f7cff', opacity: p === picked[0] ? 46 : 18 }))
        probe = { snapshot, items }
        const bytes = await snapshot.createNativePackage(items, 960, 540)
        const actual = await geometry.snapshotFromBuffer(bytes, 1, definitions)
        const selectedIndex = regions.indexOf(picked[0])
        const readback = actual.geometryForShape({ Id: selectedIndex + 2 }, 160)[0]
        assert.ok(Math.max(...readback.points.map(p => p.x)) < 491,
          'serialized native color block must also stop at cylinder wall')
      }
    }
  }
  const duplicateIntersections = g.buildGenericRegions([
    { closed: true, points: [{x:0,y:0},{x:100,y:0},{x:100,y:100},{x:0,y:100}] },
    { closed: false, points: [{x:50,y:-10},{x:50,y:110}] },
    { closed: false, points: [{x:50.0000001,y:-10},{x:50.0000001,y:110}] }
  ])
  assert.strictEqual(duplicateIntersections.length, 2, 'numerical duplicate intersections still merge')
  if (process.argv.includes('--write-probe')) {
    fs.mkdirSync('.codex-tmp/runtime-check', { recursive: true })
    fs.writeFileSync('.codex-tmp/runtime-check/can-bevel-fixed.pptx',
      await probe.snapshot.createNativePackage(probe.items, 960, 540))
  }
  console.log('Can + bevel regression passed: 20 placements/scales, distinct faces and native output')
}
main().catch(error => { console.error(error); process.exitCode = 1 })
