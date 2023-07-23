using System.Windows;
using System.Windows.Media;

namespace ScrollTabs
{
    public static class WpfExtensions
    {
        /// <summary>
        /// Checks if <paramref name="parent"/> is one of <paramref name="child"/> ancestors.
        /// </summary>
        public static bool HasParent(this IInputElement child, FrameworkElement parent)
        {
            FrameworkElement el = child as FrameworkElement;

            while (el != null)
            {
                if (el == parent)
                {
                    return true;
                }

                el = VisualTreeHelper.GetParent(el) as FrameworkElement;
            }

            return false;
        }
    }
}