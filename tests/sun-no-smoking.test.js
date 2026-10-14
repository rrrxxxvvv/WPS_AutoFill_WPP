const assert = require('assert')
const fs = require('fs')
global.JSZip = require('jszip')
const geometry = require('../js/ooxml-geometry')
require('../js/native-shape-writer')
const g = require('../js/fill-engine')._geometry

// Exact source transforms from the user's feedback-native.pptx, not screenshot
// estimates. Keep only geometry data; never overwrite the user's working deck.
const specs = [
  ['30', 'noSmoking', 2863215, 1477010, 3481070, 2842895],
  ['31', 'sun', 3129915, 768350, 4926965, 4948555]
]
async function main() {
  const definitions = fs.readFileSync('assets/presetShapeDefinitions.xml', 'utf8')
  for (const scale of [0.5, 1, 2]) {
    for (const samples of [48, 160, 320]) {
      const snapshot = await geometry.snapshotFromBuffer(fs.readFileSync('tests/fixtures/adjusted-shapes.pptx'), 1, definitions)
      specs.forEach(([id, preset, x, y, cx, cy]) => {
        snapshot.shapes[id] = { id, name: preset, preset, adjustments: {}, transform: {
          x:x*scale, y:y*scale, cx:cx*scale, cy:cy*scale, rotation:0
        } }
      })
      const paths = specs.flatMap(([Id]) => snapshot.geometryForShape({Id}, samples).filter(p => p.stroke !== false))
      const regions = g.buildGenericRegions(paths)
      assert.strictEqual(regions.length, 23, 'all intersections must split the same 23 faces at each sample density')
      const inside = {x:436*scale,y:292*scale}, outside = {x:264*scale,y:201*scale}
      const selected = regions.filter(p => g.pointInPolygon(inside,p))
      assert.strictEqual(selected.length,1)
      assert.ok(!g.pointInPolygon(outside,selected[0]), 'selected central block must not cross the sun ellipse into the left ring')
      regions.forEach(region => {
        assert.strictEqual(new Set(region.map(p => p.x.toFixed(7)+':'+p.y.toFixed(7))).size,region.length,
          'no repeated vertices / boundary retracing within a face')
        assert.ok(g.nativePaths(region)[0].commands.length < 30, 'trim original cubics, do not emit sampled line fragments')
      })
      const items = regions.map((r,i) => ({name:'TEST_'+i,paths:g.nativePaths(r),color:'#407cff',opacity:72}))
      const output = await snapshot.createNativePackage(items,960*scale,540*scale)
      const readback = await geometry.snapshotFromBuffer(output,1,definitions)
      const actual = readback.geometryForShape({Id:regions.indexOf(selected[0])+2},320)[0]
      assert.ok(g.pointInPolygon(inside,actual.points))
      assert.ok(!g.pointInPolygon(outside,actual.points), 'native output must preserve the same boundary separation')
      assert.ok(actual.curved, 'the central block must keep real cubic edges')
    }
  }
  console.log('Sun + noSmoking: 9 density/scale cases, 23 simple faces, separate central block and native curve output passed')
}
main().catch(error => {console.error(error);process.exitCode=1})
