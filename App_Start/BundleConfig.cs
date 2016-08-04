using System;
using System.Web;
using System.Web.Optimization;

namespace MVC_CRUD
{
    public class BundleConfig
    {
        // For more information on Bundling, visit http://go.microsoft.com/fwlink/?LinkId=254725
        public static void AddDefaultIgnorePatterns(IgnoreList ignoreList)
        {
            if (ignoreList == null)
                throw new ArgumentNullException("ignoreList");
            ignoreList.Ignore("*.intellisense.js");
            ignoreList.Ignore("*-vsdoc.js");
            ignoreList.Ignore("*.debug.js", OptimizationMode.WhenEnabled);
            //ignoreList.Ignore("*.min.js", OptimizationMode.WhenDisabled);
            //ignoreList.Ignore("*.min.css", OptimizationMode.WhenDisabled);
        }

        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.IgnoreList.Clear();
            AddDefaultIgnorePatterns(bundles.IgnoreList);
            /* Global */
            bundles.Add(new StyleBundle("~/Content/themes/base/css").Include(
                "~/Content/themes/base/jquery.ui.core.css",
                "~/Content/themes/base/jquery.ui.resizable.css",
                "~/Content/themes/base/jquery.ui.selectable.css",
                "~/Content/themes/base/jquery.ui.accordion.css",
                "~/Content/themes/base/jquery.ui.autocomplete.css",
                "~/Content/themes/base/jquery.ui.button.css",
                "~/Content/themes/base/jquery.ui.dialog.css",
                "~/Content/themes/base/jquery.ui.slider.css",
                "~/Content/themes/base/jquery.ui.tabs.css",
                "~/Content/themes/base/jquery.ui.datepicker.css",
                "~/Content/themes/base/jquery.ui.progressbar.css",
                "~/Content/themes/base/jquery.ui.theme.css"
            ));

            /* Bundles of Style Sheet */
            bundles.Add(new StyleBundle("~/Asset/css").Include(
                //< !--Bootstrap Plugin-- >
                "~/Asset/css/bootstrap.min.css",
                //< !--Bootstrap Material Design-- >
                "~/Asset/css/ripples.min.css",
                "~/Asset/css/app.css",
                "~/Asset/css/style.css",
                "~/Asset/css/bootstrap-datetimepicker.min.css",
                "~/Asset/css/bootstrap-select.min.css",
                //< !--Font Awesome-- >
                "~/Asset/css/font-awesome/css/font-awesome.min.css",
                //< !--Library-- >
                "~/Libs/bower/animate.css/animate.min.css",
                "~/Libs/bower/perfect-scrollbar/css/perfect-scrollbar.css",
                "~/Libs/bower/material-design-iconic-font/dist/css/material-design-iconic-font.css"
            ));

            /* Bundles of Script */
            bundles.Add(new ScriptBundle("~/Asset/js").Include(
                        "~/Asset/js/bootstrap.min.js",

                        "~/Asset/js/material.min.js",
                        "~/Asset/js/ripples.min.js",
                        "~/Scripts/modernizr-2.5.3.js",

                        "~/Libs/bower/jquery/dist/jquery.js",
                        "~/Libs/bower/jquery-ui/jquery-ui.min.js",
                        "~/Libs/bower/jQuery-Storage-API/jquery.storageapi.min.js",
                        "~/Libs/bower/bootstrap-sass/assets/javascripts/bootstrap.js",
                        "~/Libs/bower/superfish/dist/js/hoverIntent.js",
                        "~/Libs/bower/superfish/dist/js/superfish.js",
                        "~/Libs/bower/jquery-slimscroll/jquery.slimscroll.js",
                        "~/Libs/bower/perfect-scrollbar/js/perfect-scrollbar.jquery.js",
                        "~/Libs/bower/PACE/pace.min.js",
                        "~/Asset/js/bootstrap-datetimepicker.min.js",
                        "~/Asset/js/bootstrap-datetimepicker.id.js",
                        "~/Asset/js/bootstrap-select.min.js",

                        "~/Asset/js/library.js",
                        "~/Asset/js/plugins.js",
                        "~/Asset/js/app.js",
                        "~/Libs/bower/moment/moment.js",
                        "~/Libs/bower/fullcalendar/dist/fullcalendar.min.js"
            ));

        }
    }
}