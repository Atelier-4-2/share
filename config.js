/* Atelier 4·2 Share — public configuration. Both IDs are public identifiers, not secrets.
   No API key and no client secret are used anywhere.
   GOOGLE_CLIENT_ID : Google Cloud Console > APIs & Services > Credentials > OAuth client ID (Web application).
   MS_CLIENT_ID     : Azure portal > Microsoft Entra ID > App registrations > Application (client) ID (SPA).
   An empty value disables that provider's button ("설정 필요"). */
window.ATELIER_SHARE_CONFIG = {
  GOOGLE_CLIENT_ID: '333501036096-o6m50cvb01b4t49kcedi4jq4a7kv42i3.apps.googleusercontent.com',
  MS_CLIENT_ID: '',
  ROOT_FOLDER_NAME: 'Atelier 4·2 Share',
};
