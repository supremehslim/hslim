from .runner import main
from .services import ServiceError
import sys
try:
    main()
except (ServiceError, ValueError, OSError) as e:
    print('실행 중단:', e, file=sys.stderr)
    sys.exit(1)
except KeyboardInterrupt:
    print('자동화를 종료했습니다.')
