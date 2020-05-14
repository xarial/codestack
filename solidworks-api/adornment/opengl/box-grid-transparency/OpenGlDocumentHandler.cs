using CodeStack.SwEx.AddIn.Base;
using SolidWorks.Interop.sldworks;
using System.Drawing;
using static CodeStack.OpenGlBoxGrid.OpenGL;

namespace CodeStack.OpenGlBoxGrid
{
    public class OpenGlDocumentHandler : IDocumentHandler
    {
        private const int INST_COUNT = 27;
        private const int ROWS_COUNT = 3;
        private const int COLUMNS_COUNT = 3;
        private const double DIST = 0.05;
        private const double WIDTH = 0.1;
        private const double HEIGHT = 0.1;
        private const double LENGTH = 0.1;
        private readonly Color COLOR = Color.FromArgb(200, Color.Blue);

        private ModelView m_View;

        public void Init(ISldWorks app, IModelDoc2 model)
        {
            m_View = model.IActiveView;

            if (m_View != null)
            {
                m_View.BufferSwapNotify += OnBufferSwapNotify;
            }
        }

        private int OnBufferSwapNotify()
        {
            int level = 0;
            int row = 0;
            int column = 0;

            for (int i = 0; i < INST_COUNT; i++)
            {
                if (row == ROWS_COUNT)
                {
                    row = 0;
                    column++;

                    if (column == COLUMNS_COUNT)
                    {
                        column = 0;
                        level++;
                    }
                }

                RenderBox(new double[] 
                {
                    row * (WIDTH + DIST),
                    column * (LENGTH + DIST),
                    level * (HEIGHT + DIST)
                },
                WIDTH, LENGTH, HEIGHT,
                COLOR);

                row++;
            }
            
            return 0;
        }

        private void RenderBox(double[] pt, double width, double length, double height, Color color)
        {
            var vertices = new double[][]
            {
                new double[] { pt[0] - width / 2, pt[1] + length / 2, pt[2] + height / 2 },
                new double[] { pt[0] - width / 2, pt[1] - length / 2, pt[2] + height / 2 },
                new double[] { pt[0] + width / 2, pt[1] + length / 2, pt[2] + height / 2 },
                new double[] { pt[0] + width / 2, pt[1] - length / 2, pt[2] + height / 2 },
                new double[] { pt[0] + width / 2, pt[1] + length / 2, pt[2] - height / 2 },
                new double[] { pt[0] + width / 2, pt[1] - length / 2, pt[2] - height / 2 },
                new double[] { pt[0] - width / 2, pt[1] + length / 2, pt[2] - height / 2 },
                new double[] { pt[0] - width / 2, pt[1] - length / 2, pt[2] - height / 2 },
                new double[] { pt[0] - width / 2, pt[1] + length / 2, pt[2] + height / 2 },
                new double[] { pt[0] - width / 2, pt[1] - length / 2, pt[2] + height / 2 }
            };
            
            RenderTriangleStrip(vertices, color);

            RenderTriangleStrip(new double[][]
            {
                vertices[1], vertices[7], vertices[3], vertices[5]
            }, color);

            RenderTriangleStrip(new double[][]
            {
                vertices[0], vertices[2], vertices[6], vertices[4]
            }, color);
        }
        
        private void RenderTriangleStrip(double[][] vertices, Color color)
        {
            glDisable(GL_LIGHTING);
            glEnable(GL_BLEND);

            glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA);

            glBegin(GL_TRIANGLE_STRIP);

            glColor4f(color.R / 255f, color.G / 255f, color.B / 255f, color.A / 255f);

            foreach (var vertex in vertices)
            {
                glVertex3d(vertex[0], vertex[1], vertex[2]);
            }

            glEnd();
        }

        public void Dispose()
        {
            if (m_View != null)
            {
                m_View.BufferSwapNotify -= OnBufferSwapNotify;
            }
        }
    }
}
