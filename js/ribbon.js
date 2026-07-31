function OnAddinLoad(ribbonUI) {
  if (typeof window.Application.ribbonUI !== 'object') {
    window.Application.ribbonUI = ribbonUI
  }
  if (typeof window.Application.Enum !== 'object') {
    window.Application.Enum = WPS_Enum
  }
  return true
}

function OnAction(control) {
  const id = control.Id
  if (id !== 'btnOpenAutoFillPane') return true

  let paneId = window.Application.PluginStorage.getItem('wps_autofill_taskpane_id')
  let pane = null

  if (paneId) {
    try {
      pane = window.Application.GetTaskPane(paneId)
    } catch (error) {
      pane = null
    }
  }

  if (!pane) {
    pane = window.Application.CreateTaskPane(GetUrlPath() + '/ui/taskpane.html', '自动填色')
    if (!pane) {
      window.Application.alert('无法创建任务窗格，请确认当前 WPS 支持 JS 加载项。')
      return true
    }
    window.Application.PluginStorage.setItem('wps_autofill_taskpane_id', pane.ID)
    try {
      pane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
      pane.Width = 340
    } catch (error) {
      // 某些版本不允许设置宽度或停靠位置，不影响主要功能。
    }
  }

  pane.Visible = !pane.Visible
  return true
}

function GetImage(control) {
  if (control.Id === 'btnOpenAutoFillPane') return 'images/fill.svg'
  return 'images/fill.svg'
}
