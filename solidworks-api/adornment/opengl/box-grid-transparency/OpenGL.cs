using System.Runtime.InteropServices;

namespace CodeStack.OpenGlBoxGrid
{
    public static class OpenGL
    {
        [DllImport("opengl32")]
        public static extern void glBegin(uint mode);

        [DllImport("opengl32")]
        public static extern void glEnd();

        [DllImport("opengl32")]
        public static extern void glVertex3d(double x, double y, double z);

        [DllImport("opengl32.dll")]
        public static extern void glDisable(uint cap);

        [DllImport("opengl32.dll")]
        public static extern void glColor4f(float R, float G, float B, float A);

        [DllImport("opengl32.dll")]
        public static extern void glEnable(uint cap);

        [DllImport("opengl32.dll")]
        public static extern void glBlendFunc(uint sfactor, uint dfactor);

        public const int GL_TRIANGLE_STRIP = 0x0005;
        public const uint GL_LIGHTING = 0x0B50;
        public const int GL_BLEND = 0x0BE2;
        public const int GL_SRC_ALPHA = 0x0302;
        public const int GL_ONE_MINUS_SRC_ALPHA = 0x0303;
    }
}
