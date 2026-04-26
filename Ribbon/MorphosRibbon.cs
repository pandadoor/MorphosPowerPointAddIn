using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace MorphosPowerPointAddIn.Ribbon
{
    [ComVisible(true)]
    public sealed class MorphosRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly ThisAddIn _addIn;

        public MorphosRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='tabMorphos' label='Morphos'>
        <group id='grpMorphos' label='Fonts and Media'>
          <toggleButton id='btnMorphos'
                        label='Open Inspector'
                        size='large'
                        imageMso='FontDialog'
                        onAction='OnToggleTaskPane'
                        getPressed='GetPressed'
                        screentip='Toggle Morphos pane'
                        supertip='Show the Morphos-style fonts and media inspector.' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public void OnToggleTaskPane(Office.IRibbonControl control, bool pressed)
        {
            _addIn.ToggleTaskPane(pressed);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return _addIn.IsTaskPaneVisible;
        }

        internal void Invalidate()
        {
            _ribbon?.InvalidateControl("btnMorphos");
        }
    }
}
