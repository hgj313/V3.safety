export interface Theme {
  colors: {
    primary: string;
    secondary: string;
    success: string;
    warning: string;
    error: string;
    background: string;
    surface: string;
    border: string;
    shadow: string;
    text: {
      primary: string;
      secondary: string;
      disabled: string;
    };
  };
  spacing: {
    xs: string;
    sm: string;
    md: string;
    lg: string;
    xl: string;
  };
  borderRadius: {
    sm: string;
    md: string;
    lg: string;
  };
  animation: {
    duration: {
      fast: string;
      normal: string;
      slow: string;
    };
    easing: {
      easeInOut: string;
      easeOut: string;
      easeIn: string;
    };
  };
}

export const lightTheme: Theme = {
  colors: {
    primary: '#1677ff',
    secondary: '#722ed1',
    success: '#52c41a',
    warning: '#faad14',
    error: '#ff4d4f',
    background: '#ffffff',
    surface: '#fafafa',
    border: '#e8e8e8',
    shadow: '#000000',
    text: {
      primary: '#262626',
      secondary: '#8c8c8c',
      disabled: '#bfbfbf',
    },
  },
  spacing: {
    xs: '4px',
    sm: '8px',
    md: '16px',
    lg: '24px',
    xl: '32px',
  },
  borderRadius: {
    sm: '4px',
    md: '8px',
    lg: '12px',
  },
  animation: {
    duration: {
      fast: '0.15s',
      normal: '0.3s',
      slow: '0.45s',
    },
    easing: {
      easeInOut: 'cubic-bezier(0.4, 0, 0.2, 1)',
      easeOut: 'cubic-bezier(0, 0, 0.2, 1)',
      easeIn: 'cubic-bezier(0.4, 0, 1, 1)',
    },
  },
};

export const darkTheme: Theme = {
  colors: {
    primary: '#1677ff',
    secondary: '#722ed1',
    success: '#52c41a',
    warning: '#faad14',
    error: '#ff4d4f',
    background: '#141414',
    surface: '#1f1f1f',
    border: '#303030',
    shadow: '#000000',
    text: {
      primary: '#ffffff',
      secondary: '#a6a6a6',
      disabled: '#595959',
    },
  },
  spacing: {
    xs: '4px',
    sm: '8px',
    md: '16px',
    lg: '24px',
    xl: '32px',
  },
  borderRadius: {
    sm: '4px',
    md: '8px',
    lg: '12px',
  },
  animation: {
    duration: {
      fast: '0.15s',
      normal: '0.3s',
      slow: '0.45s',
    },
    easing: {
      easeInOut: 'cubic-bezier(0.4, 0, 0.2, 1)',
      easeOut: 'cubic-bezier(0, 0, 0.2, 1)',
      easeIn: 'cubic-bezier(0.4, 0, 1, 1)',
    },
  },
};

// 声明styled-components的主题类型
declare module 'styled-components' {
  export interface DefaultTheme extends Theme {}
} 