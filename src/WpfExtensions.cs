using System.Windows;
using System.Windows.Media;

namespace ScrollTabs
{
    public static class WpfExtensions
    {
        public static bool HasParent(this IInputElement child, string name)
        {
            FrameworkElement el = child as FrameworkElement;

            while (el != null)
            {
                if (el.Name == name)
                {
                    return true;
                }

                el = VisualTreeHelper.GetParent(el) as FrameworkElement;
            }

            return false;
        }
    }
}