const assert = require('assert')
const fs = require('fs')
global.JSZip = require('jszip')
global.DOMParser = require('@xmldom/xmldom').DOMParser
const writer = require('../js/native-shape-writer')
const geometry = require('../js/ooxml-geometry')
const engine = require('../js/fill-engine')
const g = engine._geometry
const probe = { name: 'EXACT_CURVE_PROBE', color: '#4f7cff', opacity: 72, paths: [{ start: { x: 100, y: 100 }, commands: [
  { kind: 'curve', p1: { x: 100, y: 200 }, p2: { x: 200, y: 200 }, end: { x: 200, y: 100 } },
  { kind: 'line', end: { x: 100, y: 100 } }
] }] }
async function main() {
  const bytes = await writer.createPackage([probe], 960, 540, fs.readFileSync('tests/fixtures/adjusted-shapes.pptx'))
  const snapshot = await geometry.snapshotFromBuffer(bytes, 1, '<presets/>')
  const paths = snapshot.geometryForShape({ Id: 2 }, 160)
  const curve = paths[0].edgeMeta.find(e => e.kind === 'curve').curve
  assert.deepStrictEqual(curve, {p0:probe.paths[0].start, p1:probe.paths[0].commands[0].p1, p2:probe.paths[0].commands[0].p2, p3:probe.paths[0].commands[0].end})
  const source = await geometry.snapshotFromBuffer(fs.readFileSync('tests/fixtures/adjusted-shapes.pptx'), 2, fs.readFileSync('assets/presetShapeDefinitions.xml', 'utf8'))
  const cases = [
    ['9004', 'snip2SameRect', 377.95001220703125, 150.9499969482422, 221.14999389648438, 238.14999389648438],
    ['9005', 'noSmoking', 236.14999389648438, 31.850000381469727, 317.54998779296875, 272.1499938964844],
    ['9012', 'heptagon', 542.3499755859375, 122.55000305175781, 283.5, 175.8000030517578]
  ]
  cases.forEach(([id,preset,x,y,cx,cy]) => { source.shapes[id]={id,name:preset,preset,adjustments:{},transform:{x:x*12700,y:y*12700,cx:cx*12700,cy:cy*12700,rotation:0},inGroup:false} })
  const sourcePaths = cases.flatMap(([id]) => source.geometryForShape({Id:Number(id)},160).filter(p=>p.stroke!==false))
  const regions = g.buildGenericRegions(sourcePaths)
  assert.strictEqual(regions.length, 9)
  const colors=['#407cff','#ffb454','#70ccac','#e992c0','#8baae8','#ab90da','#82cfdf','#ffe39a','#c1d87c']
  const items=regions.map((region,i)=>({name:'REGION_'+i,paths:g.nativePaths(region),color:colors[i],opacity:72}))
  const feedbackBytes=await source.createNativePackage(items,960,540)
  const readback=await geometry.snapshotFromBuffer(feedbackBytes,1,'<presets/>')
  items.forEach((item,i)=>{
    const actual=readback.geometryForShape({Id:i+2},160)
    const expectedCurves=item.paths.flatMap(p=>p.commands.filter(c=>c.kind==='curve'))
    const actualCurves=Array.from(new Set(actual.flatMap(p=>(p.edgeMeta||[]).filter(e=>e.kind==='curve').map(e=>e.curve))))
    assert.strictEqual(actualCurves.length,expectedCurves.length,'Native output must not re-fit/split cubic commands')
    actualCurves.forEach((c,j)=>{for(const [a,b] of [['p1','p1'],['p2','p2'],['p3','end']]){
      assert.ok(Math.hypot(c[a].x-expectedCurves[j][b].x,c[a].y-expectedCurves[j][b].y)<.0002,'Every control point must survive serialization')
    }})
  })
  // A hole must be a second subpath of ONE filled path, with reversed winding.
  const nested=g.buildGenericRegions([
    {closed:true,points:[{x:0,y:0},{x:100,y:0},{x:100,y:100},{x:0,y:100}]},
    {closed:true,points:[{x:20,y:20},{x:80,y:20},{x:80,y:80},{x:20,y:80}]}
  ])
  const outer=nested.find(p=>p._holes)
  assert.ok(outer)
  assert.ok(!g.pointInPolygon(g.polygonInteriorPoint(outer).point,outer._holes[0]),'Marker must not fall inside a hole')
  const compound=writer.shapeXml({name:'hole',paths:g.nativePaths(outer),color:'#407cff',opacity:100},0)
  assert.strictEqual((compound.match(/<a:path /g)||[]).length,1)
  assert.strictEqual((compound.match(/<a:moveTo>/g)||[]).length,2)
  // A standalone curved outer frame must retain controls after splitting.
  const oval=sourcePaths.find(p=>p.curved)
  const bounded=g.buildRegions({points:oval.points,sourcePath:oval},[],160)
  assert.ok(bounded.length && bounded.every(p=>g.freeformCommands(p)))
  let closed=false, removed=false, written=false
  const original={Id:10,Name:'untouched'}
  const slideItems=[original]
  const target={Shapes:{get Count(){return slideItems.length},Item:i=>slideItems[i-1],Paste(){
    const item={Id:11,Name:probe.name,Left:109,Top:109,Width:100,Height:100,Rotation:0,ZOrder(){},Delete(){const i=slideItems.indexOf(item);if(i>=0)slideItems.splice(i,1)}}
    slideItems.push(item)
    return {Count:1,Item:()=>item}
  }}}
  global.wps={Env:{GetTempPath:()=>'/private-test-temp'},FileSystem:{
    writeAsBinaryString(filename,data){assert.ok(filename.startsWith('/private-test-temp/WPS_AutoFill_Native_'));assert.ok(data.startsWith('PK'));written=true},
    unlinkSync(){removed=true}
  }}
  let sourceEntries=[{Name:probe.name,Left:100,Top:100,Width:100,Height:100,Rotation:0}]
  const app={Presentations:{Open(filename,readOnly,untitled,withWindow){
    assert.deepStrictEqual([readOnly,untitled,withWindow],[-1,0,0])
    return {Slides:{Item:()=>({Shapes:{Count:sourceEntries.length,Item:i=>sourceEntries[i-1],Range:()=>({Copy(){}})}})},Close(){closed=true}}
  }}}
  const imported=await writer.importPackage(app,target,bytes,[probe])
  assert.strictEqual(imported.length,1)
  assert.deepStrictEqual([imported[0].Left,imported[0].Top],[100,100],'restore paste offset without resizing')
  assert.ok(written && closed && removed)
  assert.strictEqual(slideItems[0],original)
  // A count mismatch must remove only new matching output, never original shapes.
  slideItems.splice(1)
  sourceEntries.push({...sourceEntries[0],Name:'second'})
  await assert.rejects(writer.importPackage(app,target,bytes,[probe,{...probe,name:'second'}]),/数量/)
  assert.deepStrictEqual(slideItems,[original])
  if (process.argv.includes('--write-probe')) {
    fs.mkdirSync('.codex-tmp/runtime-check', { recursive: true })
    fs.writeFileSync('.codex-tmp/runtime-check/exact-curve.pptx', bytes)
    const outlines=sourcePaths.map((p,i)=>{const polygon=p.points.slice();polygon._pathEdges=p.edgeMeta;return{name:'SOURCE_'+i,paths:g.nativePaths(polygon),color:'#204060',opacity:100,fill:false,stroke:true}})
    fs.writeFileSync('.codex-tmp/runtime-check/feedback-native.pptx',await source.createNativePackage(items.concat(outlines),960,540))
  }
  console.log('Native DrawingML control-point roundtrip passed')
}
main().catch(error => {console.error(error); process.exitCode=1})
