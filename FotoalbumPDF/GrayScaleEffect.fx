float GrayScalerRedRatio : register(C0);

float GrayScalerGreenRatio : register(C1);

float GrayScalerBlueRatio : register(C2);

sampler2D Input1 : register(S0);

float4 main(float2 uv : TEXCOORD) : COLOR
{
    float4 color;
    color = tex2D(Input1, uv);
    color.rgb = dot(color.rgb, float3(GrayScalerRedRatio, GrayScalerGreenRatio, GrayScalerBlueRatio));
    return color;
}