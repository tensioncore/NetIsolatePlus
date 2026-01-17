using System.Windows;

namespace NetIsolatePlus.Services
{
    public class WindowPlacementService
    {
        private readonly SettingsStore _store;
        private readonly string _prefix;

        public WindowPlacementService(string prefix, SettingsStore store)
        {
            _prefix = prefix;
            _store = store;
        }

        public void Restore(Window w)
        {
            var left = _store.Load($"{_prefix}.Left", double.NaN);
            var top = _store.Load($"{_prefix}.Top", double.NaN);

            if (double.IsNaN(left) || double.IsNaN(top))
            {
                var wa = SystemParameters.WorkArea;
                w.Left = wa.Left + 120; w.Top = wa.Top + 120;
                return;
            }

            var wa2 = SystemParameters.WorkArea;
            if (left < wa2.Left - 50 || top < wa2.Top - 50 ||
                left > wa2.Right - 100 || top > wa2.Bottom - 100)
            {
                w.Left = wa2.Left + 120; w.Top = wa2.Top + 120;
            }
            else
            {
                w.Left = left; w.Top = top;
            }
        }

        public void Save(Window w)
        {
            _store.Save($"{_prefix}.Left", w.Left);
            _store.Save($"{_prefix}.Top", w.Top);
        }
    }
}
