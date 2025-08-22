Here are the step-by-step instructions for installing PowerShell on MX Linux using the terminal.

### Prerequisites

First, you need to ensure your system is up to date and you have the necessary tools for the installation.

1.  Open your terminal.

2.  Run the following commands to update your package list and install the required packages:

    ```bash
    sudo apt-get update
    sudo apt-get install -y wget apt-transport-https software-properties-common
    ```

      * `wget` is a utility for downloading files from the internet.
      * `apt-transport-https` allows the package manager to use repositories over a secure connection.
      * `software-properties-common` provides tools for managing software repositories.

-----

### Step 1: Download the Microsoft Signing Key

To ensure the authenticity of the PowerShell package, you need to download and register the Microsoft GPG signing key.

1.  Use `wget` to download the key:

    ```bash
    wget -q https://packages.microsoft.com/config/debian/11/packages-microsoft-prod.deb
    ```

      * This command downloads the Debian package that contains the signing key and repository information for Debian 11, which MX Linux is based on.

-----

### Step 2: Register the Microsoft Repository

After downloading the `.deb` file, you need to install it to register the Microsoft repository and its signing key with your system.

1.  Install the package using `dpkg`:

    ```bash
    sudo dpkg -i packages-microsoft-prod.deb
    ```

      * `dpkg` is the primary tool for managing `.deb` packages on Debian-based systems. This command adds the new repository to your `/etc/apt/sources.list.d/` directory.

-----

### Step 3: Install PowerShell

With the repository now configured, you can install PowerShell directly using `apt-get`.

1.  Update your package list again to include the new repository:

    ```bash
    sudo apt-get update
    ```

2.  Install PowerShell:

    ```bash
    sudo apt-get install -y powershell
    ```

      * This command will download and install PowerShell, along with any necessary dependencies.

-----

### Step 4: Verify the Installation

Once the installation is complete, you can verify that PowerShell is installed and working correctly.

1.  Start a PowerShell session by typing `pwsh`:

    ```bash
    pwsh
    ```

2.  You should see the PowerShell prompt, which looks like this: `PS /home/your_username>`.

3.  You can then run a simple command to confirm it's working, such as checking the version:

    ```bash
    $PSVersionTable
    ```

4.  To exit the PowerShell session, simply type `exit`.

    ```bash
    exit
    ```

You now have PowerShell installed and ready to use on your MX Linux system\!
