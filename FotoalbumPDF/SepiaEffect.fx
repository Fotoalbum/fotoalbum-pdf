float GrayScalerRedRatio : register(C0);

float GrayScalerGreenRatio : register(C1);

float GrayScalerBlueRatio : register(C2);

float ColorTonerRedRatio : register(C3);

float ColorTonerGreenRatio : register(C4);

float ColorTonerBlueRatio : register(C5);

sampler2D Input1 : register(S0);

float4 main(float2 uv : TEXCOORD) : COLOR
{
    float4 color;
    color = tex2D(Input1, uv);
    color.rgb = dot(color.rgb, float3(GrayScalerRedRatio, GrayScalerGreenRatio, GrayScalerBlueRatio));
    color.r *= ColorTonerRedRatio;
    color.g *= ColorTonerGreenRatio;
    color.b *= ColorTonerBlueRatio;
    return color;
}