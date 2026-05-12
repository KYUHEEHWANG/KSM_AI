using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace KSM_SolidEdge
{
    internal class SeAppHandler
    {
    }
    public static class SolidEdgeConnector
    {
        private static SolidEdgeFramework.Application _seApp;
        public static SolidEdgeFramework.Application App
        {
            get
            {
                if (_seApp == null) _seApp = GetInstance();
                return _seApp;
            }
            set {
                _seApp = value;
            }
        }

        public static SolidEdgeFramework.Application GetInstance()
        {
            try
            {
                // 이미 연결된 객체가 있는지 확인
                if (_seApp != null)
                {
                    // 연결이 살아있는지 체크 (간단한 속성 호출로 확인)
                    var name = _seApp.Name;                    
                    return _seApp;
                }

                // 실행 중인 객체 가져오기
                _seApp = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (Exception)
            {
                // 새로 실행
                Type seType = Type.GetTypeFromProgID("SolidEdge.Application");
                _seApp = (SolidEdgeFramework.Application)Activator.CreateInstance(seType);
                _seApp.Visible = true;
            }

            return _seApp;
        }
    }
}
