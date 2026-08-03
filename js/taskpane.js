(function () {
  'use strict'

  const $ = id => document.getElementById(id)
  const controls = ['prepare', 'generateSelected', 'generate', 'recolor', 'delete']
  const STORAGE_KEY = 'WPS_AutoFill_SavedColors'
  const CLASSIC_COLORS = [
    '#000000', '#ffffff', '#e53935', '#fb8c00',
    '#fdd835', '#43a047', '#00acc1', '#1e88e5',
    '#3949ab', '#8e24aa', '#6d4c41', '#78909c'
  ]
  let savedColors = []
  let manualSelectionActive = false
  let manualPollBusy = false
  let manualDetectionText = ''

  function settings() {
    return {
      color: $('color').value,
      opacity: Number($('opacity').value),
      segments: Number($('segments').value),
      palette: $('palette').checked,
      includeNodeInteriors: $('includeNodeInteriors').checked,
      disableBoundaryFill: $('disableBoundaryFill').checked
    }
  }

  function setBusy(busy) {
    controls.forEach(id => {
      const control = $(id)
      if (control) control.disabled = busy
    })
  }

  function show(message, kind) {
    const status = $('status')
    status.className = 'status' + (kind ? ' ' + kind : '')
    status.textContent = message
  }

  function run(action) {
    setBusy(true)
    try {
      action()
    } catch (error) {
      show(error && error.message ? error.message : String(error), 'error')
    } finally {
      setBusy(false)
    }
  }

  function normalizeColor(color) {
    const value = String(color || '').trim().toLowerCase()
    return /^#[0-9a-f]{6}$/.test(value) ? value : ''
  }

  function loadSavedColors() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]')
      if (Array.isArray(parsed)) {
        savedColors = parsed.map(normalizeColor).filter(Boolean).slice(0, 8)
      }
    } catch (_) {
      savedColors = []
    }
  }

  function persistSavedColors() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(savedColors)) } catch (_) {}
  }

  function selectColor(color) {
    const normalized = normalizeColor(color)
    if (!normalized) return
    $('color').value = normalized
    updateActiveSwatches()
  }

  function createSwatch(color, removable) {
    const swatch = document.createElement('button')
    swatch.type = 'button'
    swatch.className = 'color-swatch'
    swatch.style.backgroundColor = color
    swatch.dataset.color = color
    swatch.title = removable
      ? color.toUpperCase() + '（点击使用，右键移除）'
      : color.toUpperCase()
    swatch.setAttribute('aria-label', '使用颜色 ' + color.toUpperCase())
    swatch.addEventListener('click', () => selectColor(color))
    if (removable) {
      swatch.addEventListener('contextmenu', event => {
        event.preventDefault()
        savedColors = savedColors.filter(item => item !== color)
        persistSavedColors()
        renderSavedColors()
      })
    }
    return swatch
  }

  function updateActiveSwatches() {
    const current = normalizeColor($('color').value)
    document.querySelectorAll('.color-swatch').forEach(swatch => {
      swatch.classList.toggle('active', swatch.dataset.color === current)
    })
  }

  function renderClassicColors() {
    const container = $('classicColors')
    container.textContent = ''
    CLASSIC_COLORS.forEach(color => container.appendChild(createSwatch(color, false)))
    updateActiveSwatches()
  }

  function renderSavedColors() {
    const container = $('savedColors')
    container.textContent = ''
    if (!savedColors.length) {
      const empty = document.createElement('span')
      empty.className = 'empty-colors'
      empty.textContent = '尚未保存颜色'
      container.appendChild(empty)
    } else {
      savedColors.forEach(color => container.appendChild(createSwatch(color, true)))
    }
    updateActiveSwatches()
  }

  function detectionSummary(result) {
    if (result.mode === 'network') {
      const ignored = result.ignoredLineCount
        ? '，忽略 ' + result.ignoredLineCount + ' 条未连接线段'
        : ''
      return '线网模式：' + result.faceCount + ' 个闭合面，' +
        result.nodeRegionCount + ' 个椭圆内部区域' + ignored
    }
    if (result.mode === 'generic') {
      return '通用线稿模式：读取 ' + result.lineCount + ' 条路径，识别到 ' + result.faceCount + ' 个闭合区域'
    }
    return '外框模式：识别到 ' + result.faceCount + ' 个闭合区域'
  }

  function updateManualSelectionStatus(count) {
    const button = $('generateSelected')
    if (button) button.textContent = '完成并填充（' + count + '）'
    show(
      (manualDetectionText ? manualDetectionText + '。\n' : '') +
      '已点选 ' + count + ' 个区域。直接继续点击其他区域；再次点击已加深区域可取消。完成后点击“完成并填充”。',
      count ? 'ok' : ''
    )
  }

  function stopManualSelection() {
    manualSelectionActive = false
    manualPollBusy = false
    const button = $('generateSelected')
    if (button) button.textContent = '完成并填充'
  }

  function pollManualSelection() {
    if (!manualSelectionActive || manualPollBusy) return
    manualPollBusy = true
    try {
      const result = window.WpsAutoFill.captureManualSelection(settings())
      if (result.changed) updateManualSelectionStatus(result.count)
    } catch (_) {
      // 页面切换或 WPS 正在更新选择时可能短暂无选择；下一轮继续检查。
    } finally {
      manualPollBusy = false
    }
  }

  $('opacity').addEventListener('input', event => {
    $('opacityValue').textContent = event.target.value + '%'
  })

  $('color').addEventListener('input', updateActiveSwatches)

  $('saveColor').addEventListener('click', () => {
    const color = normalizeColor($('color').value)
    if (!color) return
    savedColors = [color].concat(savedColors.filter(item => item !== color)).slice(0, 8)
    persistSavedColors()
    renderSavedColors()
    show('已临时保存颜色 ' + color.toUpperCase() + '。右键点击临时色块可单独移除。', 'ok')
  })

  $('clearSavedColors').addEventListener('click', () => {
    savedColors = []
    persistSavedColors()
    renderSavedColors()
    show('已清空临时颜色。', 'ok')
  })

  $('prepare').addEventListener('click', () => run(() => {
    stopManualSelection()
    const result = window.WpsAutoFill.prepareManualSelection(settings())
    manualDetectionText = detectionSummary(result)
    manualSelectionActive = true
    updateManualSelectionStatus(0)
  }))

  $('generateSelected').addEventListener('click', () => run(() => {
    const result = window.WpsAutoFill.generateSelected(settings())
    stopManualSelection()
    show('已生成选中的 ' + result.count + ' 个填色区域，其余候选区域已清除。', 'ok')
  }))

  $('generate').addEventListener('click', () => run(() => {
    stopManualSelection()
    const result = window.WpsAutoFill.generate(settings())
    show(
      detectionSummary(result) + '。\n已填充全部 ' + result.count + ' 个区域，并置于原轮廓和连线下方。',
      'ok'
    )
  }))

  $('recolor').addEventListener('click', () => run(() => {
    const count = window.WpsAutoFill.recolorSelected(settings())
    show('已为 ' + count + ' 个选中区域更新颜色。', 'ok')
  }))

  $('delete').addEventListener('click', () => run(() => {
    stopManualSelection()
    const count = window.WpsAutoFill.deleteGenerated()
    show(count ? '已清理本页 ' + count + ' 个候选或生成区域。' : '本页没有插件生成的区域。', 'ok')
  }))

  loadSavedColors()
  renderClassicColors()
  renderSavedColors()
  window.setInterval(pollManualSelection, 140)
})()
