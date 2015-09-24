using System;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Effects;
#if SILVERLIGHT
using UIPropertyMetadata = System.Windows.PropertyMetadata;
#endif

namespace FAEffects.Effects
{
    public class GrayScaleEffect : ShaderEffect
    {
        private static PixelShader _pixelShader = new PixelShader();

        static GrayScaleEffect()
        {
            StringBuilder uriString = new StringBuilder();
#if !SILVERLIGHT
            uriString.Append("pack://application:,,,");
#endif
            uriString.Append("/FAEffects.Effects;component/GrayScaleEffect.ps");
            _pixelShader.UriSource = new Uri(uriString.ToString(), UriKind.RelativeOrAbsolute);
        }

        public GrayScaleEffect()
        {
            this.PixelShader = _pixelShader;
            UpdateShaderValue(Input1Property);
            UpdateShaderValue(GrayScalerRedRatioProperty);
            UpdateShaderValue(GrayScalerGreenRatioProperty);
            UpdateShaderValue(GrayScalerBlueRatioProperty);
        }

        ///<summary>
        ///Gets or sets explicitly the main input sampler.
        ///</summary>
#if !SILVERLIGHT
        [System.ComponentModel.BrowsableAttribute(false)]
#endif
        public Brush Input1
        {
            get { return (Brush)GetValue(Input1Property); }
            set { SetValue(Input1Property, value); }
        }

        ///<summary>
        ///Identifies the Input1 dependency property.
        ///</summary>
        public static readonly DependencyProperty Input1Property = ShaderEffect.RegisterPixelShaderSamplerProperty("Input1", typeof(GrayScaleEffect), 0);

        public System.Double GrayScalerRedRatio
        {
            get { return (System.Double)GetValue(GrayScalerRedRatioProperty); }
            set { SetValue(GrayScalerRedRatioProperty, value); }
        }

        ///<summary>
        ///Identifies the GrayScalerRedRatio dependency property.
        ///</summary>
        public static readonly DependencyProperty GrayScalerRedRatioProperty = DependencyProperty.Register("GrayScalerRedRatio", typeof(System.Double), typeof(GrayScaleEffect), new UIPropertyMetadata(0.30, PixelShaderConstantCallback(0)));

        public System.Double GrayScalerGreenRatio
        {
            get { return (System.Double)GetValue(GrayScalerGreenRatioProperty); }
            set { SetValue(GrayScalerGreenRatioProperty, value); }
        }

        ///<summary>
        ///Identifies the GrayScalerGreenRatio dependency property.
        ///</summary>
        public static readonly DependencyProperty GrayScalerGreenRatioProperty = DependencyProperty.Register("GrayScalerGreenRatio", typeof(System.Double), typeof(GrayScaleEffect), new UIPropertyMetadata(0.59, PixelShaderConstantCallback(1)));

        public System.Double GrayScalerBlueRatio
        {
            get { return (System.Double)GetValue(GrayScalerBlueRatioProperty); }
            set { SetValue(GrayScalerBlueRatioProperty, value); }
        }

        ///<summary>
        ///Identifies the GrayScalerBlueRatio dependency property.
        ///</summary>
        public static readonly DependencyProperty GrayScalerBlueRatioProperty = DependencyProperty.Register("GrayScalerBlueRatio", typeof(System.Double), typeof(GrayScaleEffect), new UIPropertyMetadata(0.11, PixelShaderConstantCallback(2)));

    }
}