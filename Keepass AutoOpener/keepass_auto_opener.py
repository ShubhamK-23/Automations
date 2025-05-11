import os
import subprocess
import sys
import getpass
from cryptography.fernet import Fernet

class SecureConfigManager:
    def __init__(self, config_path=None):
        """
        Manage secure configuration for KeePass databases
        
        Args:
            config_path (str, optional): Path to the configuration file. 
                                         Defaults to a file in the user's home directory.
        """
        if config_path is None:
            config_path = os.path.join(os.path.expanduser('~'), '.keepass_config')
        
        self.config_path = config_path
        self.key_path = f"{config_path}.key"
        
    def _generate_key(self):
        """
        Generate a new encryption key and save it securely
        """
        key = Fernet.generate_key()
        with open(self.key_path, 'wb') as key_file:
            os.chmod(self.key_path, 0o600)  # Make file readable only by the owner
            key_file.write(key)
        return key
    
    def _load_key(self):
        """
        Load the encryption key
        """
        if not os.path.exists(self.key_path):
            return self._generate_key()
        
        with open(self.key_path, 'rb') as key_file:
            return key_file.read()
    
    def encrypt_password(self, password):
        """
        Encrypt a password
        
        Args:
            password (str): Password to encrypt
        
        Returns:
            str: Encrypted password
        """
        key = self._load_key()
        f = Fernet(key)
        return f.encrypt(password.encode()).decode()
    
    def decrypt_password(self, encrypted_password):
        """
        Decrypt a password
        
        Args:
            encrypted_password (str): Encrypted password to decrypt
        
        Returns:
            str: Decrypted password
        """
        key = self._load_key()
        f = Fernet(key)
        return f.decrypt(encrypted_password.encode()).decode()
    
    def save_database_config(self, databases):
        """
        Save database configurations securely
        
        Args:
            databases (dict): Dictionary of database paths and encrypted passwords
        """
        import json
        
        with open(self.config_path, 'w') as config_file:
            os.chmod(self.config_path, 0o600)  # Make file readable only by the owner
            json.dump(databases, config_file)
    
    def load_database_config(self):
        """
        Load database configurations
        
        Returns:
            dict: Database configurations
        """
        import json
        
        if not os.path.exists(self.config_path):
            return {}
        
        with open(self.config_path, 'r') as config_file:
            return json.load(config_file)

def setup_keepass_config():
    """
    Interactive setup for KeePass database configurations
    """
    config_manager = SecureConfigManager()
    
    # KeePass executable path
    keepass_path = os.environ.get(
        'KEEPASS_PATH', 
        r"C:\Program Files\KeePass Password Safe 2\KeePass.exe"
    )
    
    # Database configurations
    databases = {
        r"C:\Shubham Work\ViaMedici\KeePass\My DB\myDB.kdbx": "",
        r"C:\Shubham Work\ViaMedici\KeePass\KundenDB\KundenDB.kdbx": ""
    }
    
    print("KeePass Database Configuration")
    print("-----------------------------")
    
    # Collect passwords for each database
    for db_path in databases.keys():
        print(f"\nDatabase: {db_path}")
        password = getpass.getpass("Enter database password (leave blank if none): ")
        
        if password:
            # Encrypt the password
            encrypted_password = config_manager.encrypt_password(password)
            databases[db_path] = encrypted_password
    
    # Save the configuration
    config_manager.save_database_config(databases)
    
    print("\nConfiguration saved successfully!")
    return keepass_path

def open_keepass_databases():
    """
    Open KeePass databases from secure configuration
    """
    config_manager = SecureConfigManager()
    
    try:
        # Load KeePass path from environment or use default
        keepass_path = os.environ.get(
            'KEEPASS_PATH', 
            r"C:\Program Files\KeePass Password Safe 2\KeePass.exe"
        )
        
        # Load database configurations
        databases = config_manager.load_database_config()
        
        if not databases:
            print("No database configurations found. Please run setup first.")
            return
        
        # Start KeePass
        subprocess.Popen([keepass_path])
        subprocess.TimeoutExpired
        time.sleep(2)
        
        # Open each database
        for db_path, encrypted_password in databases.items():
            if os.path.exists(db_path):
                # Prepare open command
                open_cmd = [keepass_path, db_path]
                
                # Decrypt and add password if exists
                if encrypted_password:
                    decrypted_password = config_manager.decrypt_password(encrypted_password)
                    open_cmd.append(f"-pw:{decrypted_password}")
                
                # Open the database
                subprocess.Popen(open_cmd)
                time.sleep(1)
            else:
                print(f"Database not found: {db_path}")
    
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

def main():
    # Check if configuration needs setup
    config_manager = SecureConfigManager()
    if not os.path.exists(config_manager.config_path):
        print("No configuration found. Running setup...")
        setup_keepass_config()
    
    # Open databases
    open_keepass_databases()

if __name__ == "__main__":
    import time
    main()